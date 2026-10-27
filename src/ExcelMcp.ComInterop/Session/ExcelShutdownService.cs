using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Polly;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Centralized service for Excel workbook close and application quit operations.
/// Implements resilient shutdown with exponential backoff for COM busy conditions.
/// </summary>
public static class ExcelShutdownService
{
    private static readonly ResiliencePipeline _quitPipeline = ResiliencePipelines.CreateExcelQuitPipeline();

    internal static Action? BeforeQuitHookForTests { get; set; }
    internal static Action? BeforeSaveHookForTests { get; set; }

    internal static void CloseSavedWorkbookForReopen(Excel.Workbook workbook)
    {
        if (!workbook.Saved)
            throw new InvalidOperationException("The workbook has unsaved changes. Nothing was closed.");
        // Keep Excel's save prompt if a user edit races the saved-state check.
        workbook.Close();
    }

    internal static void CloseProtectedViewWindow(Excel.ProtectedViewWindow window) => window.Close();

    /// <summary>
    /// Closes an internally owned, unsaved calculation workbook on the calling STA.
    /// Unlike application shutdown, failure to close is an operation failure.
    /// The caller retains responsibility for releasing the workbook reference.
    /// </summary>
    /// <param name="workbook">Internally owned workbook; never the session workbook</param>
    public static void CloseTemporaryWorkbook(Excel.Workbook workbook)
    {
        _quitPipeline.Execute(() => workbook.Close(SaveChanges: false));
    }

    /// <summary>
    /// Saves an Excel workbook on the calling STA thread.
    /// Must be called from within <c>ExcelBatch.Execute()</c> so the Save() COM call
    /// runs on the correct STA thread. Timeout protection is provided by the surrounding
    /// <c>ExcelBatch.Execute()</c> wait deadline, which preserves outstanding Excel work.
    /// </summary>
    /// <param name="workbook">Excel workbook COM object to save</param>
    /// <param name="fileName">File name for diagnostic messages (optional)</param>
    /// <param name="logger">Logger for diagnostic output (optional)</param>
    /// <param name="cancellationToken">Cancellation token checked before Save() is invoked</param>
    /// <exception cref="OperationCanceledException">Cancellation was requested before save started</exception>
    /// <exception cref="ExcelBusyException">Excel rejected the final save because it is busy</exception>
    /// <exception cref="InvalidOperationException">Save failed due to unexpected error</exception>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Only bounded retries consume Excel's generic save error. Exhaustion reaches the enclosing COM catch, which throws with the original COM exception; no save failure becomes success.")]
    public static void SaveWorkbookWithTimeout(
        Excel.Workbook workbook,
        string? fileName = null,
        ILogger? logger = null,
        CancellationToken cancellationToken = default)
    {
        logger ??= NullLogger.Instance;
        fileName ??= "unknown";

        // Honour any cancellation request before we start the potentially slow COM call
        cancellationToken.ThrowIfCancellationRequested();

        logger.LogDebug("Saving workbook {FileName}", fileName);

        try
        {
            if (workbook.ReadOnly)
            {
                throw new InvalidOperationException(
                    $"Cannot save '{fileName}': Excel opened the workbook read-only. " +
                    "Changes have not been saved; any unsaved changes remain in the open workbook. " +
                    "Inspect the workbook's access and permissions before continuing.");
            }

            // Call Save() directly on the STA thread.
            // Previously this used Task.Run(() => workbook.Save()) which marshalled the COM call
            // from the STA thread to an MTA thread-pool thread — crossing COM apartment boundaries
            // and risking RPC_E_WRONG_THREAD or deadlocks (GitHub #482, Bug 1).
            // This method is always called from inside ExcelBatch.Execute() which already runs
            // on the dedicated STA thread, so a direct call is correct and safe.
            const int maxSaveAttempts = 3;
            const int saveRetryDelayMs = 500;

            for (int attempt = 1; attempt <= maxSaveAttempts; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    Excel.Application? application = null;
                    try
                    {
                        application = workbook.Application;
                        ExcelBusyException.ThrowIfNotReady(
                            WorkbookRefreshProbe.Read(application, workbook), "save");
                    }
                    finally
                    {
                        ComUtilities.Release(ref application);
                    }
                    BeforeSaveHookForTests?.Invoke();
                    workbook.Save();
                    if (!workbook.Saved)
                    {
                        throw new InvalidOperationException(
                            $"Workbook '{fileName}' was not saved. Excel cancelled the save or left unsaved changes. " +
                            "Changes remain in the open workbook; inspect its state before retrying.");
                    }
                    logger.LogDebug("Workbook {FileName} saved successfully", fileName);
                    return; // Success — exit method
                }
                catch (COMException ex) when (
                    attempt < maxSaveAttempts &&
                    (ex.HResult == unchecked((int)0x800A03EC) || ex.HResult == unchecked((int)0x800AC472)))
                {
                    logger.LogDebug("Save attempt {Attempt} for {FileName} got a transient Excel error (0x{HResult:X8}), retrying in {Delay}ms",
                        attempt, fileName, ex.HResult, saveRetryDelayMs * attempt);
                    Thread.Sleep(saveRetryDelayMs * attempt);
                }


                // Other COMException falls through to existing catch block below
            }
        }
        catch (COMException ex)
        {
            logger.LogError(ex, "Save failed for {FileName} (HResult: 0x{HResult:X8})", fileName, ex.HResult);
            throw CreateSaveFailureException(ex, fileName);
        }
        // All other exceptions propagate; no generic catch block.
    }

    internal static Exception CreateSaveFailureException(COMException exception, string fileName)
    {
        var message = CreateSaveFailureMessage(exception.HResult, fileName, exception.Message);
        return exception.HResult == unchecked((int)0x800AC472)
            ? new ExcelBusyException(message, exception)
            : new InvalidOperationException(message, exception);
    }

    internal static string CreateSaveFailureMessage(int hresult, string fileName, string detail) =>
        hresult switch
        {
            unchecked((int)0x800A03EC) =>
                $"Cannot save '{fileName}'. " +
                "The file may be read-only, locked by another process, or the path may not exist.",
            unchecked((int)0x800AC472) =>
                $"Cannot save '{fileName}': Excel is busy (0x800AC472). " +
                "A refresh or another Excel operation may still be running. Changes have not been saved. " +
                "Inspect the open workbook, wait for Excel to finish, then retry saving.",
            ResiliencePipelines.RPC_E_DISCONNECTED =>
                $"Cannot save '{fileName}'. Excel disconnected from automation before the save completed. " +
                "Saving has not been confirmed. Inspect the retained session and verify whether changes were saved.",
            _ => $"Failed to save workbook '{fileName}': {detail}"
        };

    /// <summary>
    /// Closes a workbook and quits Excel only when no workbooks remain.
    /// Handles save semantics, workbook close, COM object release, and resilient Quit with backoff.
    /// </summary>
    /// <param name="workbook">Excel workbook COM object (can be null)</param>
    /// <param name="excel">Excel application COM object (can be null)</param>
    /// <param name="save">True to save before closing, false to discard changes</param>
    /// <param name="filePath">File path for diagnostic logging (optional)</param>
    /// <param name="logger">Logger for diagnostic output (optional)</param>
    /// <param name="releaseWorkbook">Whether this method owns and should release the workbook reference.</param>
    /// <remarks>
    /// <para><b>Shutdown Order:</b></para>
    /// <list type="number">
    /// <item>If save=true: Call workbook.Save()</item>
    /// <item>Drain unreachable dynamic-binder metadata while Excel can answer COM releases</item>
    /// <item>Close workbook with Close(save) - save param controls Excel's save prompt</item>
    /// <item>Release the workbook COM reference when owned by this method</item>
    /// <item>Quit Excel only when no workbooks remain; otherwise leave it open</item>
    /// <item>Release the Excel COM reference</item>
    /// </list>
    /// <para><b>Resilience:</b> Retries Quit() on COM busy errors (RPC_E_SERVERCALL_RETRYLATER, RPC_E_CALL_REJECTED)</para>
    /// <para><b>Timeout:</b> No overall timeout - relies on retry exhaustion. Non-retriable errors bubble immediately.</para>
    /// </remarks>
    public static void CloseAndQuit(
        Excel.Workbook? workbook,
        Excel.Application? excel,
        bool save,
        string? filePath = null,
        ILogger? logger = null,
        bool releaseWorkbook = true) =>
        _ = CloseAndQuitAndReport(workbook, excel, save, filePath, logger, releaseWorkbook);

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Shutdown retains the original close/quit failure as the transported cause, does not continue to Quit after a failed workbook close, and retains application ownership when Quit fails. Attached workbooks pass releaseWorkbook:false so their owner releases its retained reference exactly once.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ApiDesign", "RS0030",
        Justification = "Collect unreachable dynamic-binder ITypeInfo RCWs while Excel can answer releases, without waiting for finalizers on the STA. This is supplementary to explicit release of all acquired references.")]
    internal static bool CloseAndQuitAndReport(
        Excel.Workbook? workbook,
        Excel.Application? excel,
        bool save,
        string? filePath = null,
        ILogger? logger = null,
        bool releaseWorkbook = true,
        bool? originalDisplayAlerts = null)
    {
        logger ??= NullLogger.Instance;
        string fileName = string.IsNullOrEmpty(filePath) ? "unknown" : Path.GetFileName(filePath);

        var stopwatch = Stopwatch.StartNew();
        var quitConfirmed = false;

        try
        {
            // Step 1: Explicit save if requested (before Close call)
            if (save && workbook != null)
            {
                SaveWorkbookWithTimeout(workbook, fileName, logger);
            }

            // Drain unreachable interop metadata while Excel can still service releases.
            // Acquired workbook/application references are released explicitly below.
            GC.Collect();
            Marshal.CleanupUnusedObjectsInCurrentContext();

            // Step 2: Close workbook with retry for transient COM busy errors
            if (workbook != null)
            {
                CloseWorkbookWithRetry(workbook, save, fileName, logger);
                if (excel is not null && IsWorkbookOpen(excel, workbook))
                {
                    throw new InvalidOperationException(
                        $"Excel did not close the selected workbook '{fileName}'. It remains open; Excel has not been asked to quit.");
                }

                if (releaseWorkbook)
                {
                    ComUtilities.Release(ref workbook);
                    logger.LogDebug("Released workbook COM reference for {FileName}", fileName);
                }
            }

            // Step 4: Quit Excel application with resilient retry + overall timeout
            if (excel != null)
            {
                var remainingWorkbooks = GetOpenWorkbookCount(excel);
                if (remainingWorkbooks > 0)
                {
                    logger.LogInformation(
                        "Leaving Excel open with {WorkbookCount} workbook(s) after closing {FileName}",
                        remainingWorkbooks, fileName);
                    LeaveOpenAndVisible(excel, originalDisplayAlerts);
                    ComUtilities.Release(ref excel);
                    return false;
                }

                int attemptNumber = 0;
                Exception? lastException = null;

                // This bounds retry waits, not an outstanding synchronous Quit call.
                using var quitTimeout = new CancellationTokenSource(ComInteropConstants.ExcelQuitTimeout);

                try
                {
                    logger.LogDebug("Attempting to quit Excel for {FileName} with resilient retry ({Timeout} timeout)", fileName, ComInteropConstants.ExcelQuitTimeout);

                    // Inner retry pipeline handles transient COM busy errors within the timeout
                    quitConfirmed = _quitPipeline.Execute(cancellationToken =>
                    {
                        attemptNumber++;
                        try
                        {
                            logger.LogDebug("Quit attempt {Attempt} for {FileName}", attemptNumber, fileName);
                            BeforeQuitHookForTests?.Invoke();
                            var workbookCount = GetOpenWorkbookCount(excel);
                            if (workbookCount > 0)
                            {
                                logger.LogInformation(
                                    "Leaving Excel open with {WorkbookCount} workbook(s) before Quit attempt {Attempt} for {FileName}",
                                    workbookCount, attemptNumber, fileName);
                                LeaveOpenAndVisible(excel, originalDisplayAlerts);
                                return false;
                            }
                            excel.Quit();
                            logger.LogDebug("Quit attempt {Attempt} succeeded for {FileName}", attemptNumber, fileName);
                            return true;
                        }
                        catch (COMException ex)
                        {
                            logger.LogWarning(ex,
                                "Quit attempt {Attempt} failed for {FileName} (HResult: 0x{HResult:X8})",
                                attemptNumber, fileName, ex.HResult);
                            throw; // Let pipeline decide if retry
                        }
                    }, quitTimeout.Token);

                    if (quitConfirmed)
                    {
                        logger.LogInformation(
                            "[DIAG-QUIT-SUCCESS] Excel quit succeeded for {FileName} after {Attempts} attempt(s) in {Elapsed}ms",
                            fileName, attemptNumber, stopwatch.ElapsedMilliseconds);
                        SessionDiagnostics.WriteStdErr(
                            $"[DIAG-QUIT-SUCCESS] Excel quit succeeded for {fileName} after {attemptNumber} attempt(s) in {stopwatch.ElapsedMilliseconds}ms");
                    }
                }
                catch (OperationCanceledException ex) when (quitTimeout.Token.IsCancellationRequested)
                {
                    logger.LogError(
                        "[DIAG-QUIT-TIMEOUT] Excel quit TIMED OUT after {Timeout} for {FileName} (Attempts: {Attempts}). " +
                        "The quit outcome is unconfirmed; process ownership is retained.",
                        ComInteropConstants.ExcelQuitTimeout, fileName, attemptNumber);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-QUIT-TIMEOUT] Excel quit wait expired after {ComInteropConstants.ExcelQuitTimeout} for {fileName}; process ownership is retained.");
                    lastException = new TimeoutException($"Excel.Quit() wait expired after {ComInteropConstants.ExcelQuitTimeout} for {fileName}", ex);
                }
                catch (COMException ex) when (ex.HResult == ResiliencePipelines.RPC_E_CALL_FAILED)
                {
                    logger.LogError(ex,
                        "[DIAG-QUIT-RPC-FAILED] Excel RPC connection FAILED (0x800706BE) for {FileName}. " +
                        "The quit outcome is unknown; process ownership is retained.",
                        fileName);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-QUIT-RPC-FAILED] Excel RPC connection failed for {fileName}; process ownership is retained.");
                    lastException = ex;
                }
                catch (COMException ex) when (ex.HResult == ResiliencePipelines.RPC_E_DISCONNECTED)
                {
                    logger.LogWarning(ex,
                        "Excel COM proxy disconnected while calling Quit for {FileName}; process ownership is retained",
                        fileName);
                    lastException = ex;
                }
                catch (COMException ex)
                {
                    // All retry attempts exhausted or non-retriable error
                    // INSTRUMENTATION: Track retry exhaustion path
                    logger.LogError(ex,
                        "[DIAG-QUIT-RETRY-EXHAUSTED] Excel quit failed for {FileName} after {Attempts} attempt(s) (HResult: 0x{HResult:X8}, Elapsed: {Elapsed}ms); process ownership is retained",
                        fileName, attemptNumber, ex.HResult, stopwatch.ElapsedMilliseconds);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-QUIT-RETRY-EXHAUSTED] Excel quit failed for {fileName} after {attemptNumber} attempt(s) (HResult: 0x{ex.HResult:X8}); process ownership is retained.");
                    lastException = ex;
                }
                catch (MissingMemberException ex)
                {
                    logger.LogWarning(ex,
                        "Excel Quit could not be confirmed for {FileName}; process ownership is retained",
                        fileName);
                    lastException = ex;
                }
                finally
                {
                    if (lastException is null)
                    {
                        ComUtilities.Release(ref excel);
                        logger.LogDebug("Released Excel application COM reference for {FileName}", fileName);
                    }
                }

                // Additional diagnostic if quit failed
                if (lastException != null)
                {
                    // INSTRUMENTATION: Track cleanup after failed quit
                    logger.LogWarning(
                        "[DIAG-QUIT-CLEANUP-AFTER-FAILURE] Excel quit unsuccessful for {FileName} (Elapsed: {Elapsed}s, Type: {ExceptionType}). " +
                        "Application and process ownership are retained.",
                        fileName, stopwatch.Elapsed.TotalSeconds, lastException.GetType().Name);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-QUIT-CLEANUP-AFTER-FAILURE] Excel quit unsuccessful for {fileName}; application and process ownership are retained.");
                    throw new InvalidOperationException(
                        $"The workbook close finished, but Excel quit failed for '{fileName}'. Its process may remain open.",
                        lastException);
                }
            }

            return quitConfirmed;
        }
        finally
        {
            logger.LogDebug("Excel shutdown sequence completed for {FileName} in {Elapsed}ms",
                fileName, stopwatch.ElapsedMilliseconds);
        }
    }

    /// <summary>
    /// Excel still holds workbooks this session did not close. Leave them as they are, but
    /// show Excel so the user can find them.
    /// </summary>
    private static void LeaveOpenAndVisible(Excel.Application excel, bool? originalDisplayAlerts)
    {
        if (originalDisplayAlerts is { } displayAlerts)
            excel.DisplayAlerts = displayAlerts;
        excel.Visible = true;
        if (!excel.Visible)
            throw new InvalidOperationException(
                "Excel was left open with other workbooks, but its window could not be shown. Nothing was saved or closed.");
    }

    private static int GetOpenWorkbookCount(Excel.Application excel)
    {
        Excel.Workbooks? workbooks = null;
        Excel.ProtectedViewWindows? protectedWindows = null;
        try
        {
            workbooks = excel.Workbooks;
            protectedWindows = excel.ProtectedViewWindows;
            return checked(workbooks.Count + protectedWindows.Count);
        }
        finally
        {
            ComUtilities.Release(ref protectedWindows);
            ComUtilities.Release(ref workbooks);
        }
    }

    internal static void CloseWorkbookAndConfirm(
        Excel.Workbook workbook,
        Excel.Application excel,
        bool? save,
        string filePath,
        ILogger logger)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(excel);
        ArgumentNullException.ThrowIfNull(logger);

        var fileName = Path.GetFileName(filePath);
        if (save == true)
            SaveWorkbookWithTimeout(workbook, fileName, logger);

        CloseWorkbookWithRetry(workbook, save, fileName, logger);
        if (IsWorkbookOpen(excel, workbook))
        {
            throw new InvalidOperationException(
                $"Excel did not close the selected workbook '{fileName}'. It remains open.");
        }
    }

    private static void CloseWorkbookWithRetry(
        Excel.Workbook workbook,
        bool? save,
        string fileName,
        ILogger logger)
    {
        const int maxCloseAttempts = 3;
        const int closeRetryDelayMs = 200;

        for (int attempt = 1; attempt <= maxCloseAttempts; attempt++)
        {
            try
            {
                logger.LogDebug("Closing workbook {FileName} (save={Save}, attempt {Attempt})", fileName, save, attempt);
                workbook.Close(save.HasValue ? save.Value : Type.Missing);
                logger.LogDebug("Workbook {FileName} close call returned", fileName);
                return;
            }
            catch (COMException ex) when (
                attempt < maxCloseAttempts &&
                (ex.HResult == ResiliencePipelines.RPC_E_SERVERCALL_RETRYLATER ||
                 ex.HResult == ResiliencePipelines.RPC_E_CALL_REJECTED))
            {
                logger.LogDebug("Workbook close attempt {Attempt} got transient error (0x{HResult:X8}), retrying in {Delay}ms",
                    attempt, ex.HResult, closeRetryDelayMs);
                Thread.Sleep(closeRetryDelayMs * attempt);
            }
            catch (COMException ex)
            {
                if (ex.HResult == ResiliencePipelines.RPC_E_DISCONNECTED)
                {
                    logger.LogWarning(ex,
                        "Workbook COM proxy disconnected while closing {FileName} (HResult: 0x{HResult:X8}); close outcome is unknown",
                        fileName, ex.HResult);
                }
                else
                {
                    logger.LogWarning(ex,
                        "Failed to close workbook {FileName} (HResult: 0x{HResult:X8}); Excel will not be asked to quit",
                        fileName, ex.HResult);
                }

                throw new InvalidOperationException(
                    $"Workbook close failed for '{fileName}'. Excel has not been asked to quit; the workbook may remain open.", ex);
            }
            catch (MissingMemberException ex)
            {
                logger.LogWarning(ex,
                    "Workbook Close could not be confirmed for {FileName}; Excel will not be asked to quit",
                    fileName);
                throw new InvalidOperationException(
                    $"Workbook close could not be confirmed for '{fileName}'. Excel has not been asked to quit.", ex);
            }
        }
    }

    internal static bool IsWorkbookOpen(Excel.Application excel, Excel.Workbook target)
    {
        Excel.Workbooks? workbooks = null;
        Excel.ProtectedViewWindows? protectedWindows = null;
        try
        {
            workbooks = excel.Workbooks;
            var count = workbooks.Count;
            for (var index = 1; index <= count; index++)
            {
                Excel.Workbook? workbook = null;
                try
                {
                    workbook = workbooks[index];
                    if (ReferenceEquals(workbook, target))
                        return true;
                }
                finally
                {
                    ComUtilities.Release(ref workbook);
                }
            }
            protectedWindows = excel.ProtectedViewWindows;
            for (var index = 1; index <= protectedWindows.Count; index++)
            {
                Excel.ProtectedViewWindow? window = null;
                Excel.Workbook? workbook = null;
                try
                {
                    window = protectedWindows[index];
                    workbook = window.Workbook;
                    if (ReferenceEquals(workbook, target))
                        return true;
                }
                finally
                {
                    ComUtilities.Release(ref workbook);
                    ComUtilities.Release(ref window);
                }
            }
            return false;
        }
        finally
        {
            ComUtilities.Release(ref protectedWindows);
            ComUtilities.Release(ref workbooks);
        }
    }
}

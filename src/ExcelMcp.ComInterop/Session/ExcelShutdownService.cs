using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Polly;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Centralized service for Excel workbook close and application quit operations.
/// Implements resilient shutdown with exponential backoff for COM busy conditions.
/// </summary>
public static class ExcelShutdownService
{
    private static readonly ResiliencePipeline _quitPipeline = ResiliencePipelines.CreateExcelQuitPipeline();

    /// <summary>
    /// Saves an Excel workbook with 5-minute timeout protection.
    /// Wraps the blocking Save() COM call to prevent indefinite blocking on large files.
    /// </summary>
    /// <param name="workbook">Excel workbook COM object to save</param>
    /// <param name="fileName">File name for diagnostic messages (optional)</param>
    /// <param name="logger">Logger for diagnostic output (optional)</param>
    /// <param name="cancellationToken">Cancellation token (combined with 5-minute timeout)</param>
    /// <exception cref="TimeoutException">Save exceeded 5 minutes</exception>
    /// <exception cref="COMException">Save failed due to COM error</exception>
    /// <exception cref="InvalidOperationException">Save failed due to unexpected error</exception>
    public static void SaveWorkbookWithTimeout(
        dynamic workbook,
        string? fileName = null,
        ILogger? logger = null,
        CancellationToken cancellationToken = default)
    {
        logger ??= NullLogger.Instance;
        fileName ??= "unknown";

        logger.LogDebug("Saving workbook {FileName} (5-minute timeout)", fileName);

        try
        {
            // Wrap Save() with 5-minute timeout to prevent indefinite blocking
            var saveTask = Task.Run(() => workbook.Save());

            // Create combined timeout: user's cancellation token OR 5-minute timeout
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromMinutes(5));

            if (!saveTask.Wait(TimeSpan.FromMinutes(5), timeoutCts.Token))
            {
                logger.LogError("Save operation for {FileName} timed out after 5 minutes", fileName);
                throw new TimeoutException(
                    $"Save operation for '{fileName}' exceeded 5 minutes. " +
                    "This may indicate a very large file, slow disk I/O, or antivirus interference. " +
                    "Check file size and disk performance, then retry.");
            }

            logger.LogDebug("Workbook {FileName} saved successfully", fileName);
        }
        catch (TimeoutException)
        {
            throw; // Re-throw timeout exceptions as-is
        }
        catch (COMException ex)
        {
            string errorMessage = ex.HResult switch
            {
                unchecked((int)0x800A03EC) =>
                    $"Cannot save '{fileName}'. " +
                    "The file may be read-only, locked by another process, or the path may not exist.",
                unchecked((int)0x800AC472) =>
                    $"Cannot save '{fileName}'. " +
                    "The file is locked for editing by another user or process.",
                _ => $"Failed to save workbook '{fileName}': {ex.Message}"
            };

            logger.LogError(ex, "Save failed for {FileName} (HResult: 0x{HResult:X8})", fileName, ex.HResult);
            throw new InvalidOperationException(errorMessage, ex);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Unexpected error saving {FileName}", fileName);
            throw new InvalidOperationException(
                $"Unexpected error saving workbook '{fileName}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Closes a workbook and quits the Excel application with resilient retry logic.
    /// Handles save semantics, workbook close, COM object release, and resilient Quit with backoff.
    /// </summary>
    /// <param name="workbook">Excel workbook COM object (can be null)</param>
    /// <param name="excel">Excel application COM object (can be null)</param>
    /// <param name="save">True to save before closing, false to discard changes</param>
    /// <param name="filePath">File path for diagnostic logging (optional)</param>
    /// <param name="logger">Logger for diagnostic output (optional)</param>
    /// <remarks>
    /// <para><b>Shutdown Order:</b></para>
    /// <list type="number">
    /// <item>If save=true: Call workbook.Save()</item>
    /// <item>Close workbook with Close(save) - save param controls Excel's save prompt</item>
    /// <item>Release workbook COM reference</item>
    /// <item>Quit Excel application with exponential backoff retry (6 attempts, 200ms base delay)</item>
    /// <item>Release Excel COM reference</item>
    /// <item>Force GC collection to release final COM proxies</item>
    /// </list>
    /// <para><b>Resilience:</b> Retries Quit() on COM busy errors (RPC_E_SERVERCALL_RETRYLATER, RPC_E_CALL_REJECTED)</para>
    /// <para><b>Timeout:</b> No overall timeout - relies on retry exhaustion. Non-retriable errors bubble immediately.</para>
    /// </remarks>
    public static void CloseAndQuit(
        dynamic? workbook,
        dynamic? excel,
        bool save,
        string? filePath = null,
        ILogger? logger = null)
    {
        logger ??= NullLogger.Instance;
        string fileName = string.IsNullOrEmpty(filePath) ? "unknown" : Path.GetFileName(filePath);

        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Step 1: Explicit save if requested (before Close call)
            if (save && workbook != null)
            {
                SaveWorkbookWithTimeout(workbook, fileName, logger);
            }

            // Step 2: Close workbook
            if (workbook != null)
            {
                try
                {
                    logger.LogDebug("Closing workbook {FileName} (save={Save})", fileName, save);
                    workbook.Close(save);
                    logger.LogDebug("Workbook {FileName} closed successfully", fileName);
                }
                catch (COMException ex)
                {
                    logger.LogWarning(ex,
                        "Failed to close workbook {FileName} (HResult: 0x{HResult:X8}) - continuing with cleanup",
                        fileName, ex.HResult);
                }
                catch (MissingMemberException ex)
                {
                    // COM proxy already disconnected (RPC_E_DISCONNECTED / 0x80010108)
                    logger.LogWarning(ex,
                        "Workbook COM proxy was disconnected while calling Close for {FileName} - continuing with cleanup",
                        fileName);
                }
                finally
                {
                    // Step 3: Release workbook COM reference
                    ComUtilities.Release(ref workbook!);
                }
            }

            // Step 4: Quit Excel application with resilient retry + overall timeout
            if (excel != null)
            {
                int attemptNumber = 0;
                Exception? lastException = null;

                // Outer timeout (30s) catches truly hung Excel (modal dialogs, deadlocks)
                using var quitTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));

                try
                {
                    logger.LogDebug("Attempting to quit Excel for {FileName} with resilient retry (30s timeout)", fileName);

                    // Inner retry pipeline handles transient COM busy errors within the timeout
                    _quitPipeline.Execute(cancellationToken =>
                    {
                        attemptNumber++;
                        try
                        {
                            logger.LogDebug("Quit attempt {Attempt} for {FileName}", attemptNumber, fileName);
                            excel.Quit();
                            logger.LogDebug("Quit attempt {Attempt} succeeded for {FileName}", attemptNumber, fileName);
                        }
                        catch (COMException ex)
                        {
                            lastException = ex;
                            logger.LogWarning(ex,
                                "Quit attempt {Attempt} failed for {FileName} (HResult: 0x{HResult:X8})",
                                attemptNumber, fileName, ex.HResult);
                            throw; // Let pipeline decide if retry
                        }
                    }, quitTimeout.Token);

                    logger.LogInformation("Excel quit succeeded for {FileName} after {Attempts} attempt(s) in {Elapsed}ms",
                        fileName, attemptNumber, stopwatch.ElapsedMilliseconds);
                }
                catch (OperationCanceledException) when (quitTimeout.Token.IsCancellationRequested)
                {
                    // Overall 30s timeout reached - Excel is truly hung
                    logger.LogError(
                        "Excel quit TIMED OUT after 30 seconds for {FileName} (Attempts: {Attempts}). " +
                        "Excel is likely hung (modal dialog or deadlock). Proceeding with forced COM cleanup.",
                        fileName, attemptNumber);
                    lastException = new TimeoutException($"Excel.Quit() timed out after 30 seconds for {fileName}");
                }
                catch (COMException ex)
                {
                    // All retry attempts exhausted or non-retriable error
                    logger.LogError(ex,
                        "Excel quit failed for {FileName} after {Attempts} attempt(s) (HResult: 0x{HResult:X8}, Elapsed: {Elapsed}ms) - proceeding with COM cleanup",
                        fileName, attemptNumber, ex.HResult, stopwatch.ElapsedMilliseconds);
                    lastException = ex;
                }
                catch (MissingMemberException ex)
                {
                    logger.LogWarning(ex,
                        "Excel COM proxy was disconnected while calling Quit for {FileName} - proceeding with COM cleanup",
                        fileName);
                    lastException = ex;
                }
                finally
                {
                    // Step 5: Release Excel COM reference (even if Quit failed/timed out)
                    ComUtilities.Release(ref excel!);
                }

                // Additional diagnostic if quit failed
                if (lastException != null)
                {
                    logger.LogWarning(
                        "Excel quit unsuccessful for {FileName} (Elapsed: {Elapsed}s, Type: {ExceptionType}). " +
                        "COM cleanup completed. Process may leak if Excel remains hung.",
                        fileName, stopwatch.Elapsed.TotalSeconds, lastException.GetType().Name);
                }
            }
        }
        finally
        {
            // Step 6: COM cleanup happens automatically via RCW finalizers
            // Per Microsoft docs: "RCWs can be cleaned by the CLR without additional code"
            // GC.Collect() is rarely needed and can decrease performance
            // https://learn.microsoft.com/en-us/dotnet/standard/garbage-collection/induced
            // https://learn.microsoft.com/en-us/dotnet/framework/performance/reliability-best-practices

            logger.LogDebug("Excel shutdown sequence completed for {FileName} in {Elapsed}ms",
                fileName, stopwatch.ElapsedMilliseconds);
        }
    }
}

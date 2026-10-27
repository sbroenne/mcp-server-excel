using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed partial class ExcelBatch
{
    private readonly HashSet<Excel.Workbook> _approvedMacroWorkbooks = new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<Excel.Workbook> _approvedConnectionWorkbooks = new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<Excel.Workbook> _disabledMacroWorkbooks = new(ReferenceEqualityComparer.Instance);
    private string? _securityReopenFailure;

    internal void ThrowIfWorkbookUnavailable()
    {
        if (_securityReopenFailure is { } partialState)
            throw new InvalidOperationException(partialState);
    }

    internal WorkbookNoticeStatus ReadWorkbookNotices(string? workbookId = null)
    {
        var state = GetExecutionPreflightState();
        if (state != WorkbookRefreshState.Ready)
        {
            return new(null, null, null, "unknown", null, "unknown", false,
                [state == WorkbookRefreshState.DialogOpen
                    ? DialogDetails ?? "Excel has a blocking dialog. Respond in Excel."
                    : "Excel is busy or its state cannot be confirmed. No workbook restrictions were checked."]);
        }
        return Execute((_, _) =>
        {
            var book = SelectSecurityWorkbook(workbookId);
            Excel.ProtectedViewWindow? view = null;
            Excel.Connections? connections = null;
            try
            {
                view = FindProtectedView(book);
                var protectedView = view is not null;
                var macroState = _approvedMacroWorkbooks.Contains(book) ? "enabledByRequest"
                    : _disabledMacroWorkbooks.Contains(book) || _attachedApplication is null && !_createNewFile
                        ? "disabledAtOpen" : "unknown";
                var messages = new List<string>();
                if (protectedView)
                    messages.Add("Protected View prevents editing and active content. Enabling editing can remember trust for this file; ask for explicit approval of that consequence.");
                var readOnly = book.ReadOnly;
                if (readOnly && !protectedView)
                    messages.Add("Excel reports read-only access. This alone does not identify a security notice; check file permissions, locks, or Excel's visible message.");
                if (macroState == "disabledAtOpen")
                    messages.Add(workbookId is not null
                        ? "VBA macros were disabled when this workbook opened. Use Excel to review macro permissions and run procedures; this session cannot reopen a workbook that was already open in Excel."
                        : "VBA macros were disabled when this workbook opened. Enabling them requires a separately approved reopen of a saved workbook; opening event code can then run.");
                int? connectionCount = null;
                bool? hasVba = null;
                if (!protectedView)
                {
                    hasVba = book.HasVBProject;
                    connections = book.Connections;
                    connectionCount = connections.Count;
                }
                messages.Add("Excel does not expose every notification bar through this inspection. Unknown is not proof that content is enabled. Check Excel for other notices; legacy Excel 4.0 macros can require a separate user response.");
                return new WorkbookNoticeStatus(protectedView, readOnly, hasVba, macroState,
                    connectionCount, _approvedConnectionWorkbooks.Contains(book) ? "enableRequested" : "unknown",
                    false, messages);
            }
            finally
            {
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref view);
            }
        });
    }

    internal WorkbookSecurityResult ApplyWorkbookSecurity(
        string action, bool approved, bool allowRememberedTrust, bool allowReopen, string? workbookId = null)
    {
        if (!approved)
            throw new ArgumentException("Explicit approval is required for this workbook action. Nothing was changed.");
        if (action == "enable-editing" && !allowRememberedTrust)
            throw new ArgumentException("Enabling editing can make Excel remember this file as trusted. Explicit approval of remembered trust is required; nothing was changed.");
        if (action == "enable-connections" && !allowRememberedTrust)
            throw new ArgumentException("Enabling connections can make Excel remember trust for this file and can contact data sources. Explicit approval of remembered trust is required; nothing was changed.");
        if (action == "enable-macros" && !allowReopen)
            throw new ArgumentException("Enabling macros requires separate approval to close and reopen the saved workbook. Nothing was changed.");
        if (action is not ("enable-editing" or "enable-connections" or "enable-macros"))
            throw new ArgumentException($"Unknown workbook security action '{action}'.");
        ExcelBusyException.ThrowIfNotReady(GetExecutionPreflightState(), action, this);
        return Execute((_, _) =>
        {
            var book = SelectSecurityWorkbook(workbookId);
            if (!ExcelShutdownService.IsWorkbookOpen(_excel!, book))
                throw new InvalidOperationException("The selected workbook is no longer open. Nothing was changed; discover and select the workbook again.");
            return action switch
            {
                "enable-editing" => EnableWorkbookEditing(book),
                "enable-connections" => EnableWorkbookConnections(book),
                "enable-macros" => ReopenWorkbookWithMacros(book),
                _ => throw new InvalidOperationException("Unsupported workbook security action.")
            };
        });
    }

    private Excel.Workbook SelectSecurityWorkbook(string? workbookId) =>
        workbookId is null
            ? _workbook ?? throw new InvalidOperationException(
                _securityReopenFailure ?? "The session no longer has an open workbook.")
            : _workbooks is not null && _workbooks.TryGetValue(workbookId, out var selected)
                ? selected
                : throw new InvalidOperationException("The selected workbook is no longer open. Discover workbooks again.");

    private Excel.ProtectedViewWindow? FindProtectedView(Excel.Workbook book)
    {
        Excel.ProtectedViewWindows? windows = null;
        try
        {
            windows = _excel!.ProtectedViewWindows;
            for (var index = 1; index <= windows.Count; index++)
            {
                Excel.ProtectedViewWindow? window = null;
                Excel.Workbook? candidate = null;
                try
                {
                    window = windows[index];
                    candidate = window.Workbook;
                    if (ReferenceEquals(candidate, book))
                    {
                        var result = window;
                        window = null;
                        return result;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref candidate);
                    ComUtilities.Release(ref window);
                }
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref windows);
        }
    }

    private bool FindProtectedViewAndRelease(Excel.Workbook book)
    {
        var view = FindProtectedView(book);
        try { return view is not null; }
        finally { ComUtilities.Release(ref view); }
    }

    private WorkbookSecurityResult EnableWorkbookEditing(Excel.Workbook book)
    {
        var view = FindProtectedView(book);
        Excel.Workbook? editable = null;
        try
        {
            if (view is null)
                throw new InvalidOperationException("The selected workbook is not in Protected View. Other read-only restrictions cannot be removed by this action.");
            var security = _excel!.AutomationSecurity;
            Exception? failure = null;
            var editingStarted = false;
            var editingCompleted = false;
            try
            {
                _excel.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                editingStarted = true;
                editable = view.Edit(UpdateLinks: 0);
                editingCompleted = true;
                ReplaceSecurityWorkbook(book, ref editable);
            }
            catch (Exception ex)
            {
                failure = new InvalidOperationException(editingCompleted
                    ? "Excel left Protected View, but updating the retained workbook failed. Inspect Excel and discover the selected workbook again; no changes were rolled back."
                    : editingStarted
                        ? "Enabling editing failed and Excel's current Protected View state is unknown. Inspect the selected workbook before retrying; no changes were rolled back."
                        : "Enabling editing did not start because setting VBA opening protection failed.", ex);
                throw failure;
            }
            finally
            {
                RestoreOpeningSettings(_excel, security, null, failure);
            }
            return new("enable-editing", true, true, false,
                "Excel left Protected View for the selected workbook and may remember trust for this file. External links were not updated and VBA macros were not enabled by this action.");
        }
        finally
        {
            ComUtilities.Release(ref editable);
            ComUtilities.Release(ref view);
        }
    }

    private WorkbookSecurityResult EnableWorkbookConnections(Excel.Workbook book)
    {
        if (FindProtectedViewAndRelease(book))
            throw new InvalidOperationException("The workbook is in Protected View. Enabling connections does not also approve enabling editing.");
        if (book.ReadOnly)
            throw new InvalidOperationException("The workbook is read-only. This action does not remove file-access restrictions.");
        Excel.Connections? connections = null;
        try
        {
            connections = book.Connections;
            if (connections.Count == 0)
                return new("enable-connections", false, false, false, "The selected workbook has no data connections.");
            try
            {
                book.EnableConnections();
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "Connection enabling failed. Some connections may already be enabled or refreshing; the current state is unknown. Inspect the selected workbook before retrying. No changes were rolled back.", ex);
            }
            _approvedConnectionWorkbooks.Add(book);
            return new("enable-connections", null, true, false,
                "Excel accepted enabling all data connections in this workbook. This can start refresh, contact data sources, and remember trust. Refresh completion was not established.");
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }
    }

    private WorkbookSecurityResult ReopenWorkbookWithMacros(Excel.Workbook book)
    {
        if (_attachedApplication is not null || _workbooks?.Count != 1 || !ReferenceEquals(book, _workbook))
            throw new InvalidOperationException("Macro enabling by reopen is supported only for a workbook opened by this session with no neighboring workbooks. Use Excel for an attached workbook.");
        if (_approvedMacroWorkbooks.Contains(book))
            return new("enable-macros", false, false, false, "This workbook was already reopened with macros enabled by request; it was not reopened again.");
        if (FindProtectedViewAndRelease(book))
            throw new InvalidOperationException("The workbook is in Protected View. Macro approval does not also approve enabling editing.");
        if (book.ReadOnly || !book.Saved)
            throw new InvalidOperationException("Macro enabling cannot reopen a read-only workbook or one with unsaved changes. Nothing was saved, discarded, or closed.");
        if (!book.HasVBProject)
            throw new InvalidOperationException("The selected workbook has no VBA project. Nothing was reopened.");
        if (WorkbookLocation.IsRemote(_workbookPath))
            throw new InvalidOperationException("Macro enabling by reopen is not supported for a remote workbook. Use Excel for authentication and security decisions.");
        Excel.Workbooks? books = null;
        Excel.ProtectedViewWindows? protectedWindows = null;
        try
        {
            books = _excel!.Workbooks;
            protectedWindows = _excel.ProtectedViewWindows;
            if (books.Count != 1 || protectedWindows.Count != 0)
                throw new InvalidOperationException("Another workbook is open in this Excel instance. Nothing was reopened.");
        }
        finally
        {
            ComUtilities.Release(ref protectedWindows);
            ComUtilities.Release(ref books);
        }
        var path = WorkbookLocation.Normalize(book.FullName);
        if (!string.Equals(path, WorkbookLocation.Normalize(_workbookPath), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("The workbook location changed. Inspect the session before requesting a reopen.");
        ExcelBusyException.ThrowIfNotReady(WorkbookRefreshProbe.Read(_excel!, book), "enable-macros");
        var security = _excel!.AutomationSecurity;
        var alerts = _excel.DisplayAlerts;
        Exception? failure = null;
        try
        {
            _excel.DisplayAlerts = true;
            ExcelShutdownService.CloseSavedWorkbookForReopen(book);
            if (ExcelShutdownService.IsWorkbookOpen(_excel, book))
                throw new InvalidOperationException("Excel did not close the selected workbook. It was not reopened.");
            _workbooks!.Clear();
            _approvedConnectionWorkbooks.Remove(book);
            _disabledMacroWorkbooks.Remove(book);
            _workbook = null;
            _context = null;
            _securityReopenFailure = "The original saved workbook closed, but reopening with macros has not completed. Nothing was saved or discarded by this action. Inspect Excel before continuing.";
            Excel.Workbook? closed = book;
            ComUtilities.Release(ref closed);
            _excel.AutomationSecurity = (Office.MsoAutomationSecurity)SelectAutomationSecurity(macrosApproved: true);
            books = _excel.Workbooks;
            _workbook = books.Open(path, UpdateLinks: 0, IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false)
                ?? throw new InvalidOperationException("Excel returned no reopened workbook.");
            if (!ExcelShutdownService.IsWorkbookOpen(_excel, _workbook))
            {
                _securityReopenFailure = "The original saved workbook closed, and the reopened workbook is no longer open after opening event code ran. Inspect Excel before continuing; no automatic recovery was attempted.";
                ComUtilities.Release(ref _workbook);
                throw new InvalidOperationException("Excel did not retain the reopened workbook.");
            }
            _workbooks[path] = _workbook;
            _context = new ExcelContext(_workbookPath, _excel, _workbook) { ProcessIdentity = _excelProcessIdentity };
            _approvedMacroWorkbooks.Add(_workbook);
            _securityReopenFailure = null;
            return new("enable-macros", true, false, true,
                "The saved workbook was reopened with VBA macros enabled. Opening event code may have run and may have changed the workbook. No external-link update was requested. Company restrictions can still prevent macro execution.");
        }
        catch (Exception ex)
        {
            if (_securityReopenFailure is { } partialState)
            {
                failure = new InvalidOperationException(partialState, ex);
                throw failure;
            }
            failure = ex;
            throw;
        }
        finally
        {
            ComUtilities.Release(ref books);
            RestoreOpeningSettings(_excel, security, alerts, failure);
        }
    }

    private void ReplaceSecurityWorkbook(Excel.Workbook previous, ref Excel.Workbook? replacement)
    {
        var retained = _workbooks ?? throw new InvalidOperationException("Workbook ownership is unavailable.");
        var key = retained.Single(entry => ReferenceEquals(entry.Value, previous)).Key;
        var accepted = replacement ?? throw new InvalidOperationException("Excel returned no editable workbook.");
        retained[key] = accepted;
        replacement = null;
        _disabledMacroWorkbooks.Remove(previous);
        _disabledMacroWorkbooks.Add(accepted);
        _approvedMacroWorkbooks.Remove(previous);
        _approvedConnectionWorkbooks.Remove(previous);
        try
        {
            if (ReferenceEquals(_workbook, previous))
            {
                _workbook = accepted;
                _context = null;
                _context = new ExcelContext(_workbookPath, _excel!, accepted) { ProcessIdentity = _excelProcessIdentity };
            }
        }
        finally
        {
            Excel.Workbook? released = previous;
            ComUtilities.Release(ref released);
        }
    }

    internal bool RequiresVbaApproval(Excel.Workbook book) =>
        !_approvedMacroWorkbooks.Contains(book) &&
        (_disabledMacroWorkbooks.Contains(book) || _attachedApplication is null && !_createNewFile);

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Both setting-restoration failures are collected and thrown together with the original operation failure; no cleanup failure is hidden.")]
    private static void RestoreOpeningSettings(Excel.Application app, Office.MsoAutomationSecurity security,
        bool? alerts, Exception? primary)
    {
        List<Exception> failures = [];
        try { app.AutomationSecurity = security; }
        catch (Exception ex) { failures.Add(ex); }
        if (alerts is { } displayAlerts)
        {
            try { app.DisplayAlerts = displayAlerts; }
            catch (Exception ex) { failures.Add(ex); }
        }
        if (failures.Count == 0)
            return;
        if (primary is not null)
            failures.Insert(0, primary);
        throw new AggregateException(primary is null
            ? "The workbook action completed, but restoring Excel opening settings failed; inspect the retained workbook."
            : "The workbook action failed and restoring Excel opening settings also failed; inspect the retained workbook.", failures);
    }
}

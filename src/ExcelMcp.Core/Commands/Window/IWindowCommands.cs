using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

/// <summary>Observed result of an explicit Excel window visibility change.</summary>
public sealed class WindowVisibilityResult : OperationResult
{
    /// <summary>Whether Excel's native window is visible.</summary>
    public bool IsVisible { get; set; }
    /// <summary>Whether the revealed Excel window or its blocking dialog is in the foreground.</summary>
    public bool IsForeground { get; set; }
}

/// <summary>Observed result of arranging the Excel window.</summary>
public sealed class WindowArrangeResult : OperationResult
{
    /// <summary>Whether Windows granted foreground focus after the placement was confirmed.</summary>
    public bool IsForeground { get; set; }
}

/// <summary>
/// Result containing Excel window state information.
/// </summary>
public class WindowInfoResult : OperationResult
{
    /// <summary>Whether Excel is currently visible</summary>
    public bool IsVisible { get; set; }

    /// <summary>Window state: normal, minimized, or maximized</summary>
    public string WindowState { get; set; } = string.Empty;

    /// <summary>Window left position in points</summary>
    public double Left { get; set; }

    /// <summary>Window top position in points</summary>
    public double Top { get; set; }

    /// <summary>Window width in points</summary>
    public double Width { get; set; }

    /// <summary>Window height in points</summary>
    public double Height { get; set; }

    /// <summary>Whether this is the foreground window</summary>
    public bool IsForeground { get; set; }
}

/// <summary>
/// Control Excel window visibility, position, state, status bar, and worksheet-specific views.
/// Use to show/hide Excel, bring it to front, reposition, maximize/minimize, freeze panes,
/// split panes, set zoom, and control gridlines, headings, outline symbols, and formula display.
/// Set status bar text to give users real-time feedback during operations.
///
/// VISIBILITY: 'show' reveals/restores Excel and reports whether Windows granted foreground focus. 'hide' hides Excel.
/// Both work while Excel is waiting on a blocking dialog. Confirmed blocking dialogs automatically use the same show behavior.
/// Visibility changes are reflected in session metadata (session list shows updated state).
///
/// WINDOW STATE values: 'normal', 'minimized', 'maximized'.
///
/// ARRANGE presets: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen'.
/// Arrange verifies the requested placement and reports whether Windows granted foreground focus.
/// A focus refusal does not fail a verified placement; an incorrect placement is an error.
///
/// STATUS BAR: 'set-status-bar' displays text in Excel's status bar. 'clear-status-bar' restores default.
///
/// WORKSHEET VIEW: View properties belong to a workbook window and apply to the named active worksheet.
/// 'freeze-panes' replaces existing pane boundaries with counts from the top/left,
/// without selecting a boundary cell. Page Layout view must be changed to Normal
/// in Excel before freezing; the command does not change the view mode.
/// 'set-split' creates movable panes and disables frozen panes.
/// Zoom must be between 10 and 400 percent.
/// </summary>
[ServiceCategory("Window")]
[McpTool("window", Title = "Window Management", Destructive = false, Category = "settings",
    Description = "Control Excel window visibility, position, state, status bar, and worksheet views. " +
        "VIEW: get-view reports actual topLeftCell, selectionKind, selectedRangeAddress, activeCellAddress, panes, zoom, and display settings. set-view scrolls to top_left_cell and optionally selects a cell or continuous range with select; omitted select keeps the selection. Both activate the workbook window and worksheet and leave it active; get-view is not inspection-only. set-view preserves panes and controls the last (bottom/right) pane; frozen or hidden rows/columns may adjust the actual position. Save explicitly afterwards to retain the view. freeze-panes replaces existing panes using row/column counts from the top-left, activates the worksheet, and keeps the selection; Page Layout view is rejected, so ask the user to switch to Normal view. set-split creates movable panes and removes frozen panes. set-zoom accepts 10-400 and set-display-options controls gridlines, headings, outline symbols, and formulas. " +
        "VISIBILITY: show reveals Excel and reports isVisible and isForeground (Windows may refuse focus); hide hides it. Both work while a dialog is open; dialogs are never answered or dismissed. " +
        "WINDOW STATE: normal, minimized, maximized. ARRANGE presets: left-half, right-half, top-half, bottom-half, center, full-screen; a verified placement succeeds even if Windows refuses focus.")]
[McpReadOnlyActions("get-context", "get-info")]
public interface IWindowCommands
{
    /// <summary>
    /// Reads every window belonging to the session workbook, without activation or selection.
    /// Reports native active worksheet/chart, actual range selection and active cell where available.
    /// Hidden sessions are inspected without being shown; unavailable or unsupported selections
    /// are explicit, never replaced by the application's last selected range or another workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("get-context")]
    WindowContextResult GetContext(IExcelBatch batch);

    /// <summary>
    /// Makes Excel visible and reports whether Windows granted foreground focus.
    /// Restores a minimized window to its previous normal or maximized state.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("show")]
    WindowVisibilityResult Show(IExcelBatch batch);

    /// <summary>
    /// Hides the Excel window.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("hide")]
    WindowVisibilityResult Hide(IExcelBatch batch);

    /// <summary>
    /// Brings the Excel window to the foreground without changing visibility.
    /// Restores a visible minimized window to its previous normal or maximized state.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("bring-to-front")]
    OperationResult BringToFront(IExcelBatch batch);

    /// <summary>
    /// Gets current window information (visibility, position, size, state).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("get-info")]
    WindowInfoResult GetInfo(IExcelBatch batch);

    /// <summary>
    /// Sets the window state (normal, minimized, maximized).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="windowState">Window state: 'normal', 'minimized', or 'maximized'</param>
    [ServiceAction("set-state")]
    OperationResult SetState(IExcelBatch batch, [RequiredParameter] string windowState);

    /// <summary>
    /// Sets the window position and size in points.
    /// All parameters are optional — only provided values are changed.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="left">Window left position in points</param>
    /// <param name="top">Window top position in points</param>
    /// <param name="width">Window width in points</param>
    /// <param name="height">Window height in points</param>
    [ServiceAction("set-position")]
    OperationResult SetPosition(IExcelBatch batch, double? left = null, double? top = null, double? width = null, double? height = null);

    /// <summary>
    /// Arranges the Excel window using a named preset position and verifies the resulting placement.
    /// Makes Excel visible if hidden. Returns whether Windows granted foreground focus; focus refusal
    /// does not fail a verified placement.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="preset">Preset name: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen'</param>
    [ServiceAction("arrange")]
    WindowArrangeResult Arrange(IExcelBatch batch, [RequiredParameter] string preset);

    /// <summary>
    /// Sets the Excel status bar text. The text is visible at the bottom of the Excel window.
    /// Use to give users real-time feedback about what operation is in progress.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="text">Status bar text to display (e.g. "Building PivotTable from Sales data...")</param>
    [ServiceAction("set-status-bar")]
    OperationResult SetStatusBar(IExcelBatch batch, [RequiredParameter] string text);

    /// <summary>
    /// Clears the Excel status bar, restoring the default "Ready" text.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("clear-status-bar")]
    OperationResult ClearStatusBar(IExcelBatch batch);

    /// <summary>
    /// Gets worksheet-specific view state from the workbook window.
    /// Activates the workbook window and requested worksheet without restoring the prior active view.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose view should be inspected</param>
    [ServiceAction("get-view")]
    WorksheetViewResult GetView(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Sets the last pane's scroll position and optionally selects a cell or continuous range.
    /// Activates the named worksheet and leaves it active, preserving panes, view mode, zoom,
    /// display settings, and omitted selection. Returns the actual resulting view; frozen or
    /// hidden rows and columns may adjust the position. Does not save or intentionally show Excel.
    /// Save explicitly to retain the opening view. Changes are not rolled back on failure.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose view should be changed</param>
    /// <param name="topLeftCell">One cell on the named worksheet, such as A1; controls the last bottom/right pane</param>
    /// <param name="select">Optional cell or continuous range on the named worksheet, such as A1 or A1:B3; omitted keeps selection</param>
    [ServiceAction("set-view")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Naming", "CA1716",
        Justification = "The parameter defines the requested --select CLI flag and matching select MCP/Service input; this interface is not an external extension API.")]
    WorksheetViewResult SetView(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string topLeftCell,
        string? select = null);

    /// <summary>
    /// Freezes rows above and columns left of a pane boundary.
    /// At least one of frozenRows or frozenColumns must be greater than zero.
    /// Activates the requested worksheet, replaces existing pane boundaries, and scrolls
    /// frozen directions to the start without changing the selection.
    /// Rejects Page Layout view before changing pane settings or selection; switch
    /// to Normal view in Excel and retry. Does not automatically undo partial changes
    /// if a later Excel operation fails.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose panes should be frozen</param>
    /// <param name="frozenRows">Number of rows to freeze from the top (0-1,048,575)</param>
    /// <param name="frozenColumns">Number of columns to freeze from the left (0-16,383)</param>
    [ServiceAction("freeze-panes")]
    OperationResult FreezePanes(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        int frozenRows = 0,
        int frozenColumns = 0);

    /// <summary>
    /// Unfreezes panes and removes the pane split from a worksheet view.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose panes should be unfrozen</param>
    [ServiceAction("unfreeze-panes")]
    OperationResult UnfreezePanes(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Sets movable row and column split panes and disables frozen panes.
    /// Set both values to zero to remove all splits.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose split should be changed</param>
    /// <param name="splitRows">Number of rows above the horizontal split (0-1,048,575)</param>
    /// <param name="splitColumns">Number of columns left of the vertical split (0-16,383)</param>
    [ServiceAction("set-split")]
    OperationResult SetSplit(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        int splitRows = 0,
        int splitColumns = 0);

    /// <summary>
    /// Sets worksheet zoom from 10 through 400 percent.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose zoom should be changed</param>
    /// <param name="zoom">Zoom percentage from 10 through 400</param>
    [ServiceAction("set-zoom")]
    OperationResult SetZoom(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] int zoom);

    /// <summary>
    /// Changes worksheet gridlines, row/column headings, outline symbols, or formula display.
    /// Omitted options remain unchanged.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose display options should be changed</param>
    /// <param name="showGridlines">Whether to display cell gridlines</param>
    /// <param name="showHeadings">Whether to display row and column headings</param>
    /// <param name="showOutlineSymbols">Whether to display outline level symbols</param>
    /// <param name="showFormulas">Whether to display formulas instead of their calculated values</param>
    [ServiceAction("set-display-options")]
    OperationResult SetDisplayOptions(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        bool? showGridlines = null,
        bool? showHeadings = null,
        bool? showOutlineSymbols = null,
        bool? showFormulas = null);
}

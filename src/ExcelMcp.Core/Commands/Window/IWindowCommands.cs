using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

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
/// VISIBILITY: 'show' makes Excel visible AND brings to front. 'hide' hides Excel.
/// Visibility changes are reflected in session metadata (session list shows updated state).
///
/// WINDOW STATE values: 'normal', 'minimized', 'maximized'.
///
/// ARRANGE presets: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen'.
///
/// STATUS BAR: 'set-status-bar' displays text in Excel's status bar. 'clear-status-bar' restores default.
///
/// WORKSHEET VIEW: View properties belong to a workbook window and apply to the named active worksheet.
/// 'freeze-panes' uses row and column counts above/left of the pane boundary.
/// 'set-split' creates movable panes and disables frozen panes.
/// Zoom must be between 10 and 400 percent.
/// </summary>
[ServiceCategory("window", "Window")]
[McpTool("window", Title = "Window Management", Destructive = false, Category = "settings",
    Description = "Control Excel window visibility, position, state, status bar, and worksheet-specific views. VIEW: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, and set-display-options for gridlines, headings, outline symbols, and formulas. freeze-panes uses row/column counts above and left of the pane boundary. set-split creates movable panes and disables frozen panes. Zoom range: 10-400. VISIBILITY: show makes Excel visible and brings it to front; hide hides it. WINDOW STATE: normal, minimized, maximized. ARRANGE presets: left-half, right-half, top-half, bottom-half, center, full-screen.")]
public interface IWindowCommands
{
    /// <summary>
    /// Makes the Excel window visible and brings it to the foreground.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("show")]
    OperationResult Show(IExcelBatch batch);

    /// <summary>
    /// Hides the Excel window.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("hide")]
    OperationResult Hide(IExcelBatch batch);

    /// <summary>
    /// Brings the Excel window to the foreground without changing visibility.
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
    /// Arranges the Excel window using a named preset position.
    /// Makes Excel visible if hidden.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="preset">Preset name: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen'</param>
    [ServiceAction("arrange")]
    OperationResult Arrange(IExcelBatch batch, [RequiredParameter] string preset);

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
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet whose view should be inspected</param>
    [ServiceAction("get-view")]
    WorksheetViewResult GetView(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>
    /// Freezes rows above and columns left of a pane boundary.
    /// At least one of frozenRows or frozenColumns must be greater than zero.
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

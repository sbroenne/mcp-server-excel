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
/// Control Excel window visibility, position, state, and status bar.
/// Use to show/hide Excel, bring it to front, reposition, or maximize/minimize.
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
/// </summary>
[ServiceCategory("window", "Window")]
[McpTool("window", Title = "Window Management", Destructive = false, Category = "settings",
    Description = "Control Excel window visibility, position, state, and status bar. show/hide Excel, bring to front, reposition, maximize/minimize, set status bar text. VISIBILITY: 'show' makes Excel visible AND brings to front. 'hide' hides it. WINDOW STATE: normal, minimized, maximized. ARRANGE presets: left-half, right-half, top-half, bottom-half, center, full-screen. STATUS BAR: set-status-bar displays feedback text, clear-status-bar restores default.")]
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
}

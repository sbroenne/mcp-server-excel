namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Constants for Excel COM interop operations.
/// </summary>
public static class ComInteropConstants
{
    #region Timeouts

    /// <summary>
    /// Timeout for Excel.Quit() operation (30 seconds).
    /// With DisplayAlerts=false, Excel quits quickly. This timeout catches hung scenarios.
    /// </summary>
    public static readonly TimeSpan ExcelQuitTimeout = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Timeout for STA thread join after quit.
    /// CRITICAL: Must be >= ExcelQuitTimeout to ensure Dispose() waits for CloseAndQuit() to complete.
    /// Set to ExcelQuitTimeout + 15s margin for workbook close and COM cleanup.
    /// </summary>
    public static readonly TimeSpan StaThreadJoinTimeout = ExcelQuitTimeout + TimeSpan.FromSeconds(15);

    /// <summary>
    /// Timeout for save operations (5 minutes).
    /// Large files with Power Query or Data Model may take longer to save.
    /// </summary>
    public static readonly TimeSpan SaveOperationTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Maximum wait time for session creation file lock acquisition (5 seconds).
    /// </summary>
    public static readonly TimeSpan SessionFileLockTimeout = TimeSpan.FromSeconds(5);

    #endregion

    #region Sleep Intervals

    /// <summary>
    /// Delay between file lock acquisition retries (100ms).
    /// </summary>
    public const int FileLockRetryDelayMs = 100;

    /// <summary>
    /// Delay between session lock acquisition retries (200ms).
    /// </summary>
    public const int SessionLockRetryDelayMs = 200;

    #endregion

    #region Excel File Formats

    /// <summary>
    /// Excel Open XML Workbook format code (.xlsx).
    /// XlFileFormat.xlOpenXMLWorkbook = 51
    /// </summary>
    public const int XlOpenXmlWorkbook = 51;

    /// <summary>
    /// Excel Open XML Macro-Enabled Workbook format code (.xlsm).
    /// XlFileFormat.xlOpenXMLWorkbookMacroEnabled = 52
    /// </summary>
    public const int XlOpenXmlWorkbookMacroEnabled = 52;

    #endregion
}

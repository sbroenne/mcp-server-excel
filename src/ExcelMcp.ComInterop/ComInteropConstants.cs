namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Constants for Excel COM interop operations.
/// </summary>
public static class ComInteropConstants
{
    #region Timeouts

    /// <summary>
    /// Timeout for Excel.Quit() operation (2 minutes).
    /// </summary>
    public static readonly TimeSpan ExcelQuitTimeout = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Timeout for STA thread join after quit (2.5 minutes).
    /// Allows extra time for thread cleanup after Excel process ends.
    /// </summary>
    public static readonly TimeSpan StaThreadJoinTimeout = TimeSpan.FromMinutes(2.5);

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

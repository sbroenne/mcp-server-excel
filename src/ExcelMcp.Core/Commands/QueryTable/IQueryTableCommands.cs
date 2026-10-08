using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet QueryTable lifecycle and configuration for local COM text, CSV, and legacy web imports.
/// Use powerquery for modern connectors and transformations.
/// </summary>
[ServiceCategory("QueryTable")]
[MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Partial, false,
    Evidence = "The installed Apple Events dictionary exposes QueryTables; exact command behavior is unverified.",
    ExcelApiVersion = "Excel for Mac 16.113.1; Apple Events.",
    Blocker = "exact source, refresh completion, and cleanup semantics must pass a prompt-free real-Excel fixture")]
[McpTool("querytable", Title = "QueryTable Import Operations", Destructive = true, Category = "query",
    Description = "Local Excel COM QueryTable lifecycle and configuration. Supports text and CSV imports from local files, plus legacy HTML web imports. Use powerquery for modern connectors and transformations. QueryTables do not expose Power Query M, cloud data types, workbook coauthor presence, sharing, mentions, assignments, or other Microsoft 365 service APIs.")]
[McpReadOnlyActions("list", "view", "get-refresh-status")]
public interface IQueryTableCommands
{
    /// <summary>Lists all worksheet QueryTables in the workbook.</summary>
    [ServiceAction("list")]
    QueryTableListResult List(IExcelBatch batch);

    /// <summary>Views one QueryTable and source-specific configuration.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the QueryTable</param>
    /// <param name="queryTableName">Name of the QueryTable</param>
    [MacCapability(
        MacCapabilityTier.Unsupported,
        MacImplementationStatus.Blocked,
        false,
        Evidence = "Excel for Mac 16.113.1 exposes core QueryTable properties but omits refresh period, preserve formatting, and web selection, tables, and formatting fields required by the view result.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "the native dictionary cannot return the complete public contract; select another proven tier or record bounded limitation evidence rather than returning partial success")]
    [ServiceAction("view")]
    QueryTableViewResult View(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string queryTableName);

    /// <summary>
    /// Creates and synchronously refreshes a text or CSV QueryTable.
    /// Delimiter must be one character; encoding is a Windows code page such as 65001 for UTF-8.
    /// textQualifier: double-quote, single-quote, or none.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryTableName">Name for the new QueryTable</param>
    /// <param name="sourcePath">Full path to a readable local text or CSV file</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="destinationAddress">Top-left cell for the imported data</param>
    /// <param name="delimiter">Single-character field separator; defaults to comma</param>
    /// <param name="textQualifier">Text quoting: double-quote, single-quote, or none</param>
    /// <param name="encoding">Windows code page; 65001 is UTF-8</param>
    /// <param name="hasHeaders">Whether the first row contains column headings</param>
    [MacCapability(
        MacCapabilityTier.Unsupported,
        MacImplementationStatus.Blocked,
        false,
        Evidence = "Excel for Mac 16.113.1 exposes QueryTable elements and properties but no construction command or signature for a text source and destination.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "no faithful creation route is selected; a bounded live probe or another supported tier must prove source identity, destination ownership, refresh completion, and cleanup")]
    [ServiceAction("create-text")]
    OperationResult CreateText(
        IExcelBatch batch,
        [RequiredParameter] string queryTableName,
        [RequiredParameter] string sourcePath,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string destinationAddress,
        string delimiter = ",",
        string textQualifier = "double-quote",
        int encoding = 65001,
        bool hasHeaders = true);

    /// <summary>
    /// Creates and synchronously refreshes a legacy HTML web QueryTable.
    /// selectionType: entire-page, all-tables, or specified-tables.
    /// formatting: none, rich-text, or all.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryTableName">Name for the new QueryTable</param>
    /// <param name="url">URL of the legacy HTML web source</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="destinationAddress">Top-left cell for the imported data</param>
    /// <param name="selectionType">Web selection: entire-page, all-tables, or specified-tables</param>
    /// <param name="webTables">Comma-separated table names or indices for specified-tables selection</param>
    /// <param name="formatting">Imported web formatting: none, rich-text, or all</param>
    [MacCapability(
        MacCapabilityTier.Unsupported,
        MacImplementationStatus.Blocked,
        false,
        Evidence = "Excel for Mac 16.113.1 exposes QueryTable elements and properties but no construction command or signature for a web source and destination.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "no faithful creation route is selected; a bounded live probe or another supported tier must prove source identity, destination ownership, refresh completion, and cleanup")]
    [ServiceAction("create-web")]
    OperationResult CreateWeb(
        IExcelBatch batch,
        [RequiredParameter] string queryTableName,
        [RequiredParameter] string url,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string destinationAddress,
        string selectionType = "all-tables",
        string? webTables = null,
        string formatting = "none");

    /// <summary>Updates common QueryTable refresh and formatting settings.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the QueryTable</param>
    /// <param name="queryTableName">Existing QueryTable name</param>
    /// <param name="backgroundQuery">Enable or disable background refresh</param>
    /// <param name="refreshOnFileOpen">Refresh automatically when the workbook opens</param>
    /// <param name="refreshPeriod">Automatic refresh interval in minutes; zero disables timed refresh</param>
    /// <param name="adjustColumnWidth">Resize columns to fit refreshed data</param>
    /// <param name="preserveFormatting">Preserve cell formatting when refreshing</param>
    [MacCapability(
        MacCapabilityTier.Unsupported,
        MacImplementationStatus.Blocked,
        false,
        Evidence = "Excel for Mac 16.113.1 exposes background query, refresh-on-open, and column-width settings but omits refresh period and preserve formatting required by the public mutation contract.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "the native dictionary cannot satisfy every public property variant; select another proven tier or record bounded limitation evidence rather than silently ignoring inputs")]
    [ServiceAction("set-properties")]
    OperationResult SetProperties(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string queryTableName,
        bool? backgroundQuery = null,
        bool? refreshOnFileOpen = null,
        int? refreshPeriod = null,
        bool? adjustColumnWidth = null,
        bool? preserveFormatting = null);

    /// <summary>Synchronously refreshes one QueryTable.</summary>
    [ServiceAction("refresh")]
    OperationResult Refresh(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string queryTableName);

    /// <summary>Gets the typed QueryTable.Refreshing status.</summary>
    [ServiceAction("get-refresh-status")]
    RefreshStatusResult GetRefreshStatus(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string queryTableName);

    /// <summary>Cancels an active QueryTable refresh. An idle QueryTable is reported without error.</summary>
    [ServiceAction("cancel-refresh")]
    RefreshCancellationResult CancelRefresh(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string queryTableName);

    /// <summary>Deletes one QueryTable.</summary>
    [ServiceAction("delete")]
    OperationResult Delete(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string queryTableName);
}

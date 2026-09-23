using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet QueryTable lifecycle and configuration for local COM text, CSV, and legacy web imports.
/// Use powerquery for modern connectors and transformations.
/// </summary>
[ServiceCategory("querytable", "QueryTable")]
[McpTool("querytable", Title = "QueryTable Import Operations", Destructive = true, Category = "query",
    Description = "Local Excel COM QueryTable lifecycle and configuration. Supports text and CSV imports from local files, plus legacy HTML web imports. Use powerquery for modern connectors and transformations. QueryTables do not expose Power Query M, cloud data types, workbook coauthor presence, sharing, mentions, assignments, or other Microsoft 365 service APIs.")]
public interface IQueryTableCommands
{
    /// <summary>Lists all worksheet QueryTables in the workbook.</summary>
    [ServiceAction("list")]
    QueryTableListResult List(IExcelBatch batch);

    /// <summary>Views one QueryTable and source-specific configuration.</summary>
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

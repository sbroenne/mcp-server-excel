using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete.
/// Use pivottablefield for field operations, pivottablecalc for calculated fields and layout.
///
/// BEST PRACTICE: Use 'list' before creating. Prefer 'refresh' or field modifications over delete+recreate.
/// Delete+recreate loses field configurations, filters, sorting, and custom layouts.
///
/// REFRESH: Call 'refresh' after configuring fields with pivottablefield to update the visual display.
/// This is especially important for OLAP/Data Model PivotTables where field operations
/// are structural only and don't automatically trigger a visual refresh.
///
/// CREATE OPTIONS:
/// - 'create-from-range': Use source sheet and range address for data range
/// - 'create-from-table': Use an Excel Table (ListObject) as source
/// - 'create-from-datamodel': Use a Power Pivot Data Model table as source
/// Worksheet-source creation accepts paired shareCacheWith/shareCacheWithSheet to reuse an existing cache
/// with the exact same source identity. Omit both to keep independent caches. Use share-cache to move an existing report.
/// </summary>
[ServiceCategory("PivotTable")]
[McpTool("pivottable", Title = "PivotTable Operations", Destructive = true, Category = "analysis",
    Description = "Create, refresh, delete, and change sources and cache options for PivotTables. " +
        "CREATE: create-from-range, create-from-table, or create-from-datamodel. Worksheet-source creation can reuse an existing regular cache with the exact same source via share_cache_with/share_cache_with_sheet; omit both for an independent cache. Shared calculated fields, grouping, and cache refresh affect every PivotTable using that cache. " +
        "REFRESH: call after configuring fields with pivottable_field. Prefer refresh over delete and recreate to keep field settings. TIMEOUT: 5 min for Data Model. " +
        "SHARE-CACHE: moves an existing PivotTable onto share_cache_with without recreating it. Requires matching worksheet sources and compatible fields; grouping, calculated items, and conflicting calculated fields are rejected before changes. Disconnect slicers/timelines first; they are not reconnected. A PivotTable that already shares returns changed=false. " +
        "SOURCE: set-source gives a worksheet-backed PivotTable its own cache; set-connection moves one external PivotTable to an existing workbook connection with compatible fields. Check the connection, shared cache users, and connected slicers/timelines first and disconnect those controls; set-connection rejects shared caches and Data Model conversion. " +
        "Failures report the partial state; changes are not rolled back. Use pivottable_calc set-layout-options for native styles and preserveFormatting, and pivottable_field for fields.")]
[McpReadOnlyActions("list", "read", "get-source", "get-cache-options", "get-connection")]
public interface IPivotTableCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all PivotTables in workbook with structure details.
    /// Each PivotTable's layoutError reports a detected #SPILL! layout failure even when the list succeeds.
    /// cacheIndex is Excel's actual workbook-local cache index, not a permanent identifier.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>List of PivotTables with names, sheets, ranges, source data, field counts, last refresh</returns>
    [ServiceAction("list")]
    PivotTableListResult List(IExcelBatch batch);

    /// <summary>
    /// Gets PivotTable configuration, current layout, and a bounded semantic snapshot
    /// of selected members and native label/value filters.
    /// pivotTable.layoutError reports a detected #SPILL! layout failure even when the read succeeds.
    /// pivotTable.cacheIndex is Excel's actual workbook-local cache index, not a permanent identifier.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="sheetName">Exact worksheet containing the selected PivotTable. Optional for read; required for share-cache, get-connection, and set-connection.</param>
    /// <param name="maxSelectedMembers">Maximum member names returned across all fields (1 through 1000; default 100).</param>
    /// <returns>All fields with positions, aggregation functions, filter states</returns>
    [ServiceAction("read")]
    PivotTableInfoResult Read(
        IExcelBatch batch,
        string pivotTableName,
        string? sheetName = null,
        int maxSelectedMembers = PivotTableSnapshotHelpers.DefaultMaxSelectedMembers);

    /// <summary>
    /// Creates PivotTable from Excel range with auto-detection of headers
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name</param>
    /// <param name="sourceRange">Source range address (e.g., "A1:F100"). Cache sharing requires an explicit fixed A1 address, not a named or structured reference; use create-from-table for a Table source.</param>
    /// <param name="destinationSheet">Destination worksheet name</param>
    /// <param name="destinationCell">Destination cell address (e.g., "A1")</param>
    /// <param name="pivotTableName">Name for the new PivotTable</param>
    /// <param name="shareCacheWith">Existing worksheet-backed PivotTable whose exact source and native cache to reuse. Requires shareCacheWithSheet. Creation: omit both for an independent cache. share-cache: both are required.</param>
    /// <param name="shareCacheWithSheet">Exact worksheet containing shareCacheWith. Both inputs must be nonblank when supplied. Shared calculations and refresh affect other cache users; moving existing reports rejects grouping and calculated items. Slicer connections are separate.</param>
    /// <returns>Created PivotTable name and initial field list</returns>
    [ServiceAction("create-from-range")]
    PivotTableCreateResult CreateFromRange(IExcelBatch batch,
        string sourceSheet, string sourceRange,
        string destinationSheet, string destinationCell,
        string pivotTableName, string? shareCacheWith = null, string? shareCacheWithSheet = null);

    /// <summary>
    /// Creates PivotTable from Excel Table (ListObject)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Excel Table</param>
    /// <param name="destinationSheet">Destination worksheet name</param>
    /// <param name="destinationCell">Destination cell address (e.g., "A1")</param>
    /// <param name="pivotTableName">Name for the new PivotTable</param>
    /// <param name="shareCacheWith">Existing worksheet-backed PivotTable whose exact table source and native cache to reuse. Requires shareCacheWithSheet; omitted creates an independent cache.</param>
    /// <param name="shareCacheWithSheet">Exact worksheet containing shareCacheWith. Both sharing inputs must be nonblank. Shared calculated fields, grouping and cache refresh affect other cache users; slicer connections are separate.</param>
    /// <returns>Created PivotTable name and available fields</returns>
    [ServiceAction("create-from-table")]
    PivotTableCreateResult CreateFromTable(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName, string? shareCacheWith = null, string? shareCacheWithSheet = null);

    /// <summary>
    /// Creates PivotTable from Power Pivot Data Model table
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Data Model table</param>
    /// <param name="destinationSheet">Destination worksheet name</param>
    /// <param name="destinationCell">Destination cell address (e.g., "A1")</param>
    /// <param name="pivotTableName">Name for the new PivotTable</param>
    /// <returns>Created PivotTable name and available fields</returns>
    [ServiceAction("create-from-datamodel")]
    PivotTableCreateResult CreateFromDataModel(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName);

    /// <summary>
    /// Deletes PivotTable completely
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to delete</param>
    /// <returns>Operation result</returns>
    [ServiceAction("delete")]
    OperationResult Delete(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Refreshes PivotTable data from source and returns updated info
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to refresh</param>
    /// <param name="timeout">Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch</param>
    /// <returns>Refresh timestamp, record count, any structural changes</returns>
    [ServiceAction("refresh")]
    PivotTableRefreshResult Refresh(IExcelBatch batch, string pivotTableName, TimeSpan? timeout = null);

    /// <summary>Reads native worksheet-backed source, record count, shared cache users, and connected slicer/timeline caches. External/OLAP sources are unsupported here.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="pivotTableName">Selected regular PivotTable.</param>
    [ServiceAction("get-source")]
    PivotSourceResult GetSource(IExcelBatch batch, string pivotTableName);

    /// <summary>Changes only the selected regular PivotTable to a new worksheet-backed cache and refreshes it. Preserves field schema and unrelated shared-cache users. Connected slicers/timelines must first be disconnected; no dashboard cache is silently rebuilt.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="pivotTableName">Selected worksheet-backed PivotTable.</param>
    /// <param name="sourceSheetName">New source worksheet.</param>
    /// <param name="sourceRangeAddress">One contiguous header/data range with the same field names; mutually exclusive with tableName.</param>
    /// <param name="tableName">Existing Excel table on sourceSheetName; mutually exclusive with sourceRangeAddress.</param>
    [ServiceAction("set-source")]
    PivotSourceResult SetSource(IExcelBatch batch, string pivotTableName, string sourceSheetName,
        string? sourceRangeAddress = null, string? tableName = null);

    /// <summary>Moves an existing worksheet-backed report onto another existing report's cache without recreating reports or creating a cache. Requires identical source identity and compatible fields and cached items. Preserves ordinary report settings; rejects grouping, calculated items, and missing/conflicting calculated fields before writes. Disconnect controls from the moved report first. The chosen cache's records, calculations and refresh settings are adopted; controls are not rebuilt or reconnected. Already sharing is an unchanged outcome. Failures do not undo changes.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="sheetName">Exact worksheet containing the report to move.</param>
    /// <param name="pivotTableName">Existing report to move, without recreating it.</param>
    /// <param name="shareCacheWith">Existing worksheet-backed report whose native cache to use.</param>
    /// <param name="shareCacheWithSheet">Exact worksheet containing shareCacheWith.</param>
    [ServiceAction("share-cache")]
    PivotCacheSharingResult ShareCache(IExcelBatch batch, string sheetName, string pivotTableName,
        string shareCacheWith, string shareCacheWithSheet);

    /// <summary>Reads the selected PivotTable's connection identity, cache, shared users and connected controls without returning a connection string.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="sheetName">Worksheet containing the PivotTable; required because names can repeat across worksheets.</param>
    /// <param name="pivotTableName">Exact PivotTable name on sheetName.</param>
    [ServiceAction("get-connection")]
    PivotConnectionResult GetConnection(IExcelBatch batch, string sheetName, string pivotTableName);

    /// <summary>Uses Excel's native ChangeConnection to redirect one external PivotTable to an existing OLEDB or ODBC connection without recreating it. The target must expose compatible fields/measures. Rejects worksheet/Data Model sources, shared caches and connected slicers/timelines before writes. Does not delete connections or explicitly refresh; Excel may contact the provider during the change. Inspect get-connection after a failure; changes are not rolled back.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="sheetName">Worksheet containing the PivotTable; required because names can repeat across worksheets.</param>
    /// <param name="pivotTableName">Exact external PivotTable name on sheetName.</param>
    /// <param name="connectionName">Existing workbook connection of the same type and OLAP mode, with compatible fields/measures.</param>
    /// <param name="timeout">Optional cancellation timeout in whole seconds from 1 through 2147483; defaults to five minutes. The session operation timeout may be shorter.</param>
    [ServiceAction("set-connection")]
    PivotConnectionResult SetConnection(IExcelBatch batch, string sheetName, string pivotTableName,
        string connectionName, TimeSpan? timeout = null);

    /// <summary>
    /// Gets refresh, retention, and source-data settings for a PivotTable and its PivotCache.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <returns>Current PivotCache and source-data settings</returns>
    [ServiceAction("get-cache-options")]
    PivotCacheOptionsResult GetCacheOptions(IExcelBatch batch, string pivotTableName);

    /// <summary>
    /// Configures refresh, retention, and source-data settings for a PivotTable and its PivotCache.
    /// Omitted parameters keep their current values. Missing-item retention applies only to regular
    /// PivotTables; OLAP/Data Model caches manage retained members at the source.
    /// Cache-wide changes are rejected when other PivotTables share the cache; saveSourceData is table-specific.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="enableRefresh">Allow the PivotCache to refresh</param>
    /// <param name="refreshOnFileOpen">Refresh the PivotCache when the workbook opens</param>
    /// <param name="missingItemsLimit">How many deleted source items Excel retains in the cache</param>
    /// <param name="optimizeCache">Optimize the cache when it is constructed</param>
    /// <param name="saveSourceData">Save source records with the PivotTable</param>
    /// <returns>The applied cache options</returns>
    [ServiceAction("set-cache-options")]
    PivotCacheOptionsResult SetCacheOptions(
        IExcelBatch batch,
        string pivotTableName,
        bool? enableRefresh = null,
        bool? refreshOnFileOpen = null,
        [FromString] PivotMissingItemsLimit? missingItemsLimit = null,
        bool? optimizeCache = null,
        bool? saveSourceData = null);

    /// <summary>
    /// Expands a regular PivotTable value cell into a new worksheet containing its underlying
    /// source records. The address must identify a cell inside the PivotTable data body.
    /// OLAP/Data Model drill-through is provider-dependent and is not supported by this action.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="cellAddress">Value-cell address on the PivotTable worksheet (for example, G4)</param>
    /// <returns>The created detail worksheet and row count</returns>
    [ServiceAction("drill-through")]
    PivotDrillThroughResult DrillThrough(
        IExcelBatch batch,
        string pivotTableName,
        string cellAddress);
}

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>
/// Slicer visual filters for PivotTables and Excel Tables.
///
/// PIVOTTABLE SLICERS: create-slicer, list-slicers, set-slicer-selection, delete-slicer.
/// TABLE SLICERS: create-table-slicer, list-table-slicers, set-table-slicer-selection, delete-table-slicer.
///
/// NAMING: Caller supplies a unique name, destination worksheet, and anchor cell.
///
/// SELECTION: selectedItems as list of strings. Data Model/OLAP slicers return item captions
/// and accept captions or MDX unique names; unknown or ambiguous items fail before changing the filter.
/// Empty list clears filter (shows all items). Set clearFirst=false to add to existing selection.
/// </summary>
[ServiceCategory("Slicer")]
[McpTool("slicer", Title = "Slicer Operations", Destructive = true, Category = "analysis",
    Description = "Create, configure, and delete slicers and timelines that filter PivotTables and Tables. Creation requires a unique slicer_name, an existing destination_sheet, and a position. " +
        "PIVOTTABLE SLICERS: create-slicer, set-slicer-selection, delete-slicer; they filter only connected PivotTables. TABLE SLICERS: create-table-slicer, set-table-slicer-selection, delete-table-slicer. TIMELINES: create-timeline, set-timeline-selection, clear-timeline-selection, using calendar dates. " +
        "SELECTION: selected_items is JSON-array text; '[]' clears the filter. Selection replaces by default; clear_first=false adds. Data Model/OLAP selection accepts captions or MDX unique names and rejects unknown or ambiguous items before changing the filter. " +
        "connect-pivottable links compatible PivotTables without rebuilding caches: regular PivotTables must share the PivotCache; Data Model PivotTables must use the same workbook Data Model connection. disconnect-pivottable removes one link. update-slicer changes appearance settings in points. " +
        "STYLES: create-style copies source_style_name into a new style_name, overriding supplied slicer_style_options without applying it. apply-style changes only the named slicer's style, not layout or filters. Timeline styles use workbook style commands and update-slicer.")]
[McpReadOnlyActions("get-slicer", "list-slicers", "list-table-slicers", "list-styles")]
public interface ISlicerCommands
{
    /// <summary>Lists all built-in and custom styles available for ordinary PivotTable, Data Model, and Table slicers in this workbook. Excludes styles not available for ordinary slicers.</summary>
    [ServiceAction("list-styles")]
    TableStyleListResult ListStyles(IExcelBatch batch);

    /// <summary>Copies an existing ordinary slicer style into a new workbook-owned custom style and applies only supplied element overrides. Does not apply it to any slicer. Existing destination names are rejected. Failures after copying leave the new style for inspection; no automatic rollback.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="styleName">New workbook style name; must not already exist.</param>
    /// <param name="sourceStyleName">Existing built-in or custom ordinary slicer style from list-styles.</param>
    /// <param name="slicerStyleOptions">Typed object with optional wholeTable, header, selectedItemWithData, selectedItemWithNoData, unselectedItemWithData, unselectedItemWithNoData and four hovered equivalents. Each accepts fillColor, borderColor, fontColor (#RRGGBB) and bold. borderColor sets four solid thin outer borders. Omitted values inherit the source; an empty object copies unchanged.</param>
    [ServiceAction("create-style")]
    TableStyleResult CreateStyle(IExcelBatch batch, [RequiredParameter] string styleName,
        [RequiredParameter] string sourceStyleName, [RequiredParameter] SlicerStyleOptions slicerStyleOptions);

    /// <summary>Applies an existing built-in or custom ordinary slicer style to one PivotTable, Data Model, or Table slicer. Leaves column count, layout, connections, and filters unchanged. Timelines are not supported by this action.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="slicerName">Existing ordinary slicer name.</param>
    /// <param name="styleName">Existing ordinary slicer style from list-styles.</param>
    [ServiceAction("apply-style")]
    SlicerStateResult ApplyStyle(IExcelBatch batch, [RequiredParameter] string slicerName,
        [RequiredParameter] string styleName);

    /// <summary>Creates a native date timeline for a PivotTable date field. Does not rebuild the source or alter existing caches. Use get-slicer for native date/layout state and delete-slicer for its lifecycle.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Source PivotTable</param>
    /// <param name="fieldName">Native date field</param>
    /// <param name="slicerName">Unique timeline name</param>
    /// <param name="destinationSheet">Destination worksheet</param>
    /// <param name="position">Single top-left anchor cell</param>
    [ServiceAction("create-timeline")]
    SlicerStateResult CreateTimeline(IExcelBatch batch, [RequiredParameter] string pivotTableName,
        [RequiredParameter] string fieldName, [RequiredParameter] string slicerName,
        [RequiredParameter] string destinationSheet, [RequiredParameter] string position);

    /// <summary>Reads complete native layout, cache connections, selection and timeline state for one PivotTable/Table slicer or date timeline.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Visual control name</param>
    [ServiceAction("get-slicer")]
    SlicerStateResult GetSlicer(IExcelBatch batch, [RequiredParameter] string slicerName);

    /// <summary>Updates only supplied native layout/style settings. Timeline view settings apply only to timelines; columns/header layout settings apply only to ordinary slicers. Coordinates and dimensions use points.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Visual control name</param>
    /// <param name="slicerOptions">Typed native layout settings; nested JSON keys use camelCase</param>
    [ServiceAction("update-slicer")]
    SlicerStateResult UpdateSlicer(IExcelBatch batch, [RequiredParameter] string slicerName, [RequiredParameter] SlicerUpdateOptions slicerOptions);

    /// <summary>Sets an inclusive native timeline date range, filtering every connected PivotTable. Dates are calendar dates; time components are rejected.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Timeline name, not an ordinary slicer</param>
    /// <param name="timelineSelection">Required startDate and endDate calendar dates, in ascending order</param>
    [ServiceAction("set-timeline-selection")]
    SlicerStateResult SetTimelineSelection(IExcelBatch batch, [RequiredParameter] string slicerName, [RequiredParameter] TimelineSelectionOptions timelineSelection);

    /// <summary>Clears this timeline's native date filter on every connected PivotTable, without clearing unrelated field filters.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Timeline name</param>
    [ServiceAction("clear-timeline-selection")]
    SlicerStateResult ClearTimelineSelection(IExcelBatch batch, [RequiredParameter] string slicerName);

    /// <summary>Connects a compatible PivotTable without rebuilding caches. Regular PivotTables must share the control's existing native PivotCache; Data Model PivotTables may use separate caches when both use the same workbook Data Model connection. Shared selection affects every connected PivotTable.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">PivotTable slicer or timeline</param>
    /// <param name="pivotTableName">Compatible PivotTable to connect</param>
    [ServiceAction("connect-pivottable")]
    SlicerStateResult ConnectPivotTable(IExcelBatch batch, [RequiredParameter] string slicerName, [RequiredParameter] string pivotTableName);

    /// <summary>Disconnects one currently connected PivotTable; keeps at least one PivotTable connection. Does not reset other connected tables or their filters.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">PivotTable slicer or timeline</param>
    /// <param name="pivotTableName">Connected PivotTable to disconnect</param>
    [ServiceAction("disconnect-pivottable")]
    SlicerStateResult DisconnectPivotTable(IExcelBatch batch, [RequiredParameter] string slicerName, [RequiredParameter] string pivotTableName);

    /// <summary>
    /// Creates a slicer for a PivotTable field.
    /// Slicers provide visual filtering for PivotTable data.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable to create slicer for</param>
    /// <param name="fieldName">Name of the field to use for the slicer; Data Model/OLAP uses its discovered hierarchy name, e.g. [Quarters].[Quarter]</param>
    /// <param name="slicerName">Name for the new slicer</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed</param>
    /// <param name="position">Top-left cell position for the slicer (e.g., "H2")</param>
    /// <returns>Created slicer details with available items</returns>
    [ServiceAction("create-slicer")]
    SlicerResult CreateSlicer(IExcelBatch batch, string pivotTableName,
        string fieldName, string slicerName, string destinationSheet, string position);

    /// <summary>
    /// Lists all slicers in the workbook, optionally filtered by PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Optional PivotTable name to filter slicers (null = all slicers)</param>
    /// <returns>List of slicers with names, fields, positions, and selections</returns>
    [ServiceAction("list-slicers")]
    SlicerListResult ListSlicers(IExcelBatch batch, string? pivotTableName = null);

    /// <summary>
    /// Sets the selection for a slicer, filtering the connected PivotTable(s).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to modify</param>
    /// <param name="selectedItems">Items to select (show in PivotTable); empty clears the filter. Data Model/OLAP accepts returned captions or MDX unique names; unknown or ambiguous items fail without changing the filter.</param>
    /// <param name="clearFirst">If true, clears existing selection before setting new items (default: true)</param>
    /// <returns>Updated slicer state with current selection</returns>
    [ServiceAction("set-slicer-selection")]
    SlicerResult SetSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true);

    /// <summary>
    /// Deletes a slicer from the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to delete</param>
    /// <returns>Operation result indicating success or failure</returns>
    [ServiceAction("delete-slicer")]
    OperationResult DeleteSlicer(IExcelBatch batch, string slicerName);

    /// <summary>
    /// Creates a slicer for an Excel Table column.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Excel Table</param>
    /// <param name="columnName">Name of the column to use for the slicer</param>
    /// <param name="slicerName">Name for the new slicer</param>
    /// <param name="destinationSheet">Worksheet where slicer will be placed</param>
    /// <param name="position">Top-left cell position for the slicer (e.g., "H2")</param>
    /// <returns>Created slicer details with available items</returns>
    [ServiceAction("create-table-slicer")]
    SlicerResult CreateTableSlicer(IExcelBatch batch, string tableName,
        string columnName, string slicerName, string destinationSheet, string position);

    /// <summary>
    /// Lists all table slicers in the workbook, optionally filtered by table.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Optional table name to filter slicers (null = all table slicers)</param>
    /// <returns>List of slicers with names, columns, positions, and selections</returns>
    [ServiceAction("list-table-slicers")]
    SlicerListResult ListTableSlicers(IExcelBatch batch, string? tableName = null);

    /// <summary>
    /// Sets the selection for a table slicer, filtering the connected table.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to modify</param>
    /// <param name="selectedItems">Items to select (show in table)</param>
    /// <param name="clearFirst">If true, clears existing selection before setting new items (default: true)</param>
    /// <returns>Updated slicer state with current selection</returns>
    [ServiceAction("set-table-slicer-selection")]
    SlicerResult SetTableSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true);

    /// <summary>
    /// Deletes a table slicer from the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="slicerName">Name of the slicer to delete</param>
    /// <returns>Operation result indicating success or failure</returns>
    [ServiceAction("delete-table-slicer")]
    OperationResult DeleteTableSlicer(IExcelBatch batch, string slicerName);
}

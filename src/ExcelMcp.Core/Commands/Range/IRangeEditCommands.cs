using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Commands.Filtering;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data.
/// Use range for values/formulas/copy/clear operations.
///
/// INSERT/DELETE CELLS: Specify shift direction to control how surrounding cells move.
/// - Insert: 'Down' or 'Right'
/// - Delete: 'Up' or 'Left'
///
/// INSERT/DELETE ROWS: Use row range like '5:10' to insert/delete rows 5-10.
/// INSERT/DELETE COLUMNS: Use column range like 'B:D' to insert/delete columns B-D.
///
/// FIND/REPLACE: Search within the specified range with optional case/cell matching.
/// - Find returns up to maxMatches cells (default: 10), with exact totalCount, returnedCount, and truncated.
/// - Exact counting searches all matches, even after the returned-cell limit is reached.
/// - Replace modifies all matches by default.
///
/// SORT: Specify sortColumns as array of {columnIndex: 1, ascending: true} objects.
/// Column indices are 1-based relative to the range.
/// </summary>
[ServiceCategory("rangeedit", "RangeEdit")]
[McpTool("range_edit", Title = "Range Edit Operations", Destructive = true, Category = "data",
    Description = "Insert/delete cells, rows, or columns; replace text; sort and clean data. " +
        "remove-duplicates retains first records using explicit key_columns relative to the selected rectangle and has_headers, with exact blank-aware counts. Removed rows are cleared inside the selection, not shifted from below. " +
        "text-to-columns parses one source column using native delimited/fixed-width options. destination_cell is one same-sheet anchor. Preflight protects the complete output including empty fields; source cells may be replaced in place, other occupied cells require overwrite_policy='allow'. Nested options use camelCase. " +
        "FILTERS: apply-filter uses filter_options for native comparisons, AND/OR, values, date groups, top/bottom, colors, icons, and dynamic criteria; column_index is relative to an ordinary header-plus-data rectangle. Use table_column for tables. An active advanced row filter must be explicitly cleared before apply-filter; do not retry or clear it without authorization. clear-filters preserves dropdowns and unrelated AutoFilters; clear_advanced explicitly authorizes worksheet-wide advanced row-filter clearing when original criteria/scope cannot be read. " +
        "advanced-filter uses criteria_range for InPlace or Copy mode, optionally unique_only. Copy preflight conservatively protects the full maximum output extent and returns checkedDestinationRange, not an actual result count. " +
        "fill copies the source edge's contents and formatting through an unmerged rectangle. auto-fill uses native Excel patterns; destination_range includes source_range and extends it in exactly one direction on the same worksheet. create-series uses native DataSeries with step_value and optional stop_value; orientation selects rows or columns. " +
        "Fill/series content destinations default to overwrite_policy='reject-nonempty', excluding source edges; formats-only AutoFill preserves content. Trend fitting can replace source values and requires allow. Stop-value preflight conservatively protects the entire selected extent. Never automatically retry with allow. Excel protection still applies. " +
        "Cell movement uses insertShift (Down/Right) or deleteShift (Up/Left). Rows use a range like 5:10; columns use B:D. Replace modifies all matches by default (replace_options.replaceAll=true). Sort uses sortColumns, an array of {columnIndex, ascending}; indices are 1-based relative to the range.")]
[McpReadOnlyActions("get-filters", "find")]
public interface IRangeEditCommands
{
    /// <summary>Uses native criteria-range filtering in place or copies matches, optionally unique.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing source and output</param>
    /// <param name="rangeAddress">Complete ordinary source rectangle including headers</param>
    /// <param name="criteriaRange">Native criteria rectangle including its header row</param>
    /// <param name="mode">InPlace hides nonmatches; Copy writes matching records to copyToRange</param>
    /// <param name="copyToRange">Copy-mode single-cell anchor or one-row selected output headers</param>
    /// <param name="uniqueOnly">Retain only native unique records</param>
    /// <param name="overwritePolicy">Copy preflight protects the full maximum output extent; allow permits authorized replacement</param>
    [ServiceAction("advanced-filter")]
    AdvancedFilterResult AdvancedFilter(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress,
        [RequiredParameter] string criteriaRange, [RequiredParameter][FromString] AdvancedFilterMode mode,
        string? copyToRange = null, bool uniqueOnly = false,
        [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    /// <summary>Applies a typed native filter to an ordinary rectangle, preserving other field filters.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Exact unmerged header-plus-data rectangle; must not intersect a table or replace an active advanced row filter</param>
    /// <param name="columnIndex">One-based column relative to the rectangle</param>
    /// <param name="filterOptions">Shared typed native filter options with camelCase nested keys</param>
    [ServiceAction("apply-filter")]
    OperationResult ApplyFilter(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress,
        [RequiredParameter] int columnIndex, [RequiredParameter] FilterOptions filterOptions);

    /// <summary>Reads all columns and both native criteria, preserving arrays and reporting getter failures.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Exact ordinary filter rectangle, or intended rectangle when not yet enabled</param>
    [ServiceAction("get-filters")]
    RangeFilterResult GetFilters(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>Clears active criteria only on the exact selected ordinary filter, retaining its dropdowns.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Exact ordinary filter rectangle; unrelated table/range filters are not changed</param>
    /// <param name="clearAdvanced">Explicit permission to clear worksheet-wide advanced row filtering when its original scope cannot be inspected; default false</param>
    [ServiceAction("clear-filters")]
    OperationResult ClearFilters(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, bool clearAdvanced = false);

    /// <summary>
    /// Removes duplicate rows through Excel, retaining the first record for each key.
    /// Clears removed rows inside the selection without shifting cells below it.
    /// Counts include blank records and exclude the optional header.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Complete unmerged rectangular records, including the optional header</param>
    /// <param name="keyColumns">Distinct one-based columns relative to the rectangle; at least one and at most 16383</param>
    /// <param name="hasHeaders">Whether the first row is a header; default true, never guessed</param>
    [ServiceAction("remove-duplicates")]
    RemoveDuplicatesResult RemoveDuplicates(IExcelBatch batch, string sheetName,
        [RequiredParameter] string rangeAddress, [RequiredParameter] List<int> keyColumns, bool hasHeaders = true);

    /// <summary>
    /// Splits one source column with native TextToColumns. Native preflight establishes
    /// the entire output, including empty fields. Existing source cells may be replaced
    /// when parsing in place; other occupied destination cells require explicit allow.
    /// Temporary native calculation data is never saved or retained.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Source and destination worksheet</param>
    /// <param name="sourceRange">Single unmerged source column, with no header inference</param>
    /// <param name="destinationCell">Single unmerged output anchor; may be the source's first cell</param>
    /// <param name="options">Native delimiters, qualifiers, field types/positions, and number separators</param>
    /// <param name="overwritePolicy">reject-nonempty protects cells outside the source; allow permits authorized replacement</param>
    [ServiceAction("text-to-columns")]
    TextToColumnsResult TextToColumns(IExcelBatch batch, string sheetName,
        [RequiredParameter] string sourceRange, [RequiredParameter] string destinationCell,
        [RequiredParameter] TextToColumnsOptions options,
        [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    /// <summary>
    /// Uses native FillDown/FillUp/FillLeft/FillRight to copy the source edge's contents
    /// and formatting through one unmerged rectangle. Source edge is not a destination.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Complete unmerged rectangle, including the source edge</param>
    /// <param name="direction">down from top row, up from bottom row, left from right column, right from left column</param>
    /// <param name="overwritePolicy">reject-nonempty (default) protects destination content; allow permits authorized replacement</param>
    [ServiceAction("fill")]
    OperationResult Fill(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress,
        [RequiredParameter][FromString] FillDirection direction,
        [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    /// <summary>
    /// Extends a source pattern using native AutoFill. Destination includes the source and
    /// extends it in exactly one direction on the same worksheet. Formats-only leaves contents intact.
    /// Trend modes can also replace source values and require allow.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="sourceRange">Unmerged source rectangle containing pattern seeds</param>
    /// <param name="destinationRange">Unmerged complete destination, including the unchanged source extent</param>
    /// <param name="fillType">Native AutoFill kind; default lets Excel infer the pattern</param>
    /// <param name="overwritePolicy">reject-nonempty (default) protects new destinations; trend modes require allow for replacing source values</param>
    [ServiceAction("auto-fill")]
    OperationResult AutoFill(IExcelBatch batch, string sheetName,
        [RequiredParameter] string sourceRange, [RequiredParameter] string destinationRange,
        [FromString] AutoFillKind fillType = AutoFillKind.Default,
        [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    /// <summary>
    /// Creates native DataSeries across rows or down columns. Each series starts at its
    /// first cell. Trend fitting uses existing values and requires allow.
    /// Stop can leave part of the selected range unchanged; overwrite preflight conservatively
    /// covers the entire selected extent outside the seed edge.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Complete unmerged series rectangle, including initial values</param>
    /// <param name="orientation">rows across columns or columns down rows</param>
    /// <param name="seriesType">linear (default), growth, date, or autoFill</param>
    /// <param name="stepValue">Finite nonzero native step; default 1</param>
    /// <param name="stopValue">Optional finite stopping value; omitted fills the selected extent</param>
    /// <param name="dateUnit">Date series unit; default day</param>
    /// <param name="trend">Fit existing values for linear/growth series; requires overwritePolicy allow</param>
    /// <param name="overwritePolicy">reject-nonempty (default) protects the full possible destination; allow permits authorized replacement</param>
    [ServiceAction("create-series")]
    OperationResult CreateSeries(IExcelBatch batch, string sheetName,
        [RequiredParameter] string rangeAddress, [RequiredParameter][FromString] SeriesOrientation orientation,
        [FromString] SeriesKind seriesType = SeriesKind.Linear, double stepValue = 1,
        double? stopValue = null, [FromString] SeriesDateUnit dateUnit = SeriesDateUnit.Day,
        bool trend = false, [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    // === INSERT/DELETE CELL OPERATIONS ===

    /// <summary>
    /// Inserts blank cells, shifting existing cells down or right.
    /// Excel COM: Range.Insert(shift)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address where cells will be inserted (e.g., 'A1:D10')</param>
    /// <param name="insertShift">Direction to shift existing cells: 'Down' or 'Right'</param>
    [ServiceAction("insert-cells")]
    OperationResult InsertCells(
        IExcelBatch batch, string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter]
        [FromString] InsertShiftDirection insertShift);

    /// <summary>
    /// Deletes cells, shifting remaining cells up or left.
    /// Excel COM: Range.Delete(shift)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to delete (e.g., 'A1:D10')</param>
    /// <param name="deleteShift">Direction to shift remaining cells: 'Up' or 'Left'</param>
    [ServiceAction("delete-cells")]
    OperationResult DeleteCells(
        IExcelBatch batch, string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter]
        [FromString] DeleteShiftDirection deleteShift);

    /// <summary>
    /// Inserts entire rows above the range.
    /// Excel COM: Range.EntireRow.Insert()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row range defining rows to insert above (e.g., '5:10' for rows 5-10)</param>
    [ServiceAction("insert-rows")]
    OperationResult InsertRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire rows in the range.
    /// Excel COM: Range.EntireRow.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row range defining rows to delete (e.g., '5:10' for rows 5-10)</param>
    [ServiceAction("delete-rows")]
    OperationResult DeleteRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Inserts entire columns to the left of the range.
    /// Excel COM: Range.EntireColumn.Insert()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Column range defining columns to insert left of (e.g., 'B:D' for columns B-D)</param>
    [ServiceAction("insert-columns")]
    OperationResult InsertColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire columns in the range.
    /// Excel COM: Range.EntireColumn.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Column range defining columns to delete (e.g., 'B:D' for columns B-D)</param>
    [ServiceAction("delete-columns")]
    OperationResult DeleteColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === FIND/REPLACE OPERATIONS ===

    /// <summary>
    /// Counts all matching cells and returns up to maxMatches cell details.
    /// Find returns exact totalCount, returnedCount, and truncated with optional case/cell matching.
    /// maxMatches (default: 10) bounds returned cell details, not search time.
    /// Excel COM: Range.Find()/FindNext(). Exact counting traverses every match.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to search within (e.g., 'A1:D100')</param>
    /// <param name="searchValue">Text or value to search for</param>
    /// <param name="findOptions">Search options: matchCase (default: false), matchEntireCell (default: false), searchFormulas (default: true)</param>
    /// <param name="maxMatches">Maximum matching cells to return (default: 10). Positive whole number from 1 through 2147483647. All matches are still counted.</param>
    [ServiceAction("find")]
    RangeFindResult Find(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string searchValue, FindOptions findOptions, int maxMatches = 10);

    /// <summary>
    /// Replaces text/values in range.
    /// Excel COM: Range.Replace()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to search within (e.g., 'A1:D100')</param>
    /// <param name="findValue">Text or value to search for</param>
    /// <param name="replaceValue">Text or value to replace matches with</param>
    /// <param name="replaceOptions">Replace options: matchCase (default: false), matchEntireCell (default: false), replaceAll (default: true)</param>
    [ServiceAction("replace")]
    OperationResult Replace(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string findValue, [RequiredParameter] string replaceValue, ReplaceOptions replaceOptions);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts range by one or more columns.
    /// Excel COM: Range.Sort()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to sort (e.g., 'A1:D100')</param>
    /// <param name="sortColumns">Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range</param>
    /// <param name="hasHeaders">Whether the range has a header row to exclude from sorting (default: true)</param>
    [ServiceAction("sort")]
    OperationResult Sort(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] List<SortColumn> sortColumns, bool hasHeaders = true);
}

using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Strategy interface for PivotTable field operations.
/// Handles different PivotTable types (Regular vs OLAP/Data Model).
/// </summary>
public interface IPivotTableFieldStrategy
{
    /// <summary>
    /// Determines if this strategy can handle the given PivotTable
    /// </summary>
    bool CanHandle(dynamic pivot);

    /// <summary>
    /// Gets a field for manipulation from the PivotTable.
    /// Returns CubeField for OLAP, PivotField for regular.
    /// </summary>
    dynamic GetFieldForManipulation(dynamic pivot, string fieldName);

    /// <summary>
    /// Lists all fields in the PivotTable
    /// </summary>
    PivotFieldListResult ListFields(dynamic pivot, string workbookPath);

    /// <summary>
    /// Adds a field to the Row area
    /// </summary>
    PivotFieldResult AddRowField(dynamic pivot, string fieldName, int? position, string workbookPath);

    /// <summary>
    /// Adds a field to the Column area
    /// </summary>
    PivotFieldResult AddColumnField(dynamic pivot, string fieldName, int? position, string workbookPath);

    /// <summary>
    /// Adds a field to the Values area with aggregation
    /// </summary>
    PivotFieldResult AddValueField(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string? customName, string workbookPath);

    /// <summary>
    /// Adds a field to the Filter area
    /// </summary>
    PivotFieldResult AddFilterField(dynamic pivot, string fieldName, string workbookPath);

    /// <summary>
    /// Removes a field from any area
    /// </summary>
    PivotFieldResult RemoveField(dynamic pivot, string fieldName, string workbookPath);

    /// <summary>
    /// Sets custom name for a field
    /// </summary>
    PivotFieldResult SetFieldName(dynamic pivot, string fieldName, string customName, string workbookPath);

    /// <summary>
    /// Sets aggregation function for a value field
    /// </summary>
    PivotFieldResult SetFieldFunction(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string workbookPath);

    /// <summary>
    /// Sets format for a value field
    /// </summary>
    PivotFieldResult SetFieldFormat(dynamic pivot, string fieldName, string numberFormat, string workbookPath);

    /// <summary>
    /// Sets filter for a field
    /// </summary>
    PivotFieldFilterResult SetFieldFilter(dynamic pivot, string fieldName, List<string> filterValues, string workbookPath);

    /// <summary>
    /// Sorts a field
    /// </summary>
    PivotFieldResult SortField(dynamic pivot, string fieldName, SortDirection direction, string workbookPath);
}

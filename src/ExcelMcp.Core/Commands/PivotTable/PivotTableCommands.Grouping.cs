using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable grouping operations (GroupByDate, GroupByNumeric)
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Groups a date/time field by the specified interval (Month, Quarter, Year)
    /// </summary>
    public PivotFieldResult GroupByDate(IExcelBatch batch, string pivotTableName,
        string fieldName, DateGroupingInterval interval)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.GroupByDate(pivot, fieldName, interval, batch.WorkbookPath, batch.Logger));

    /// <summary>
    /// Groups a numeric field by the specified interval (e.g., 0-100, 100-200, 200-300)
    /// </summary>
    public PivotFieldResult GroupByNumeric(IExcelBatch batch, string pivotTableName,
        string fieldName, double? start, double? endValue, double intervalSize)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.GroupByNumeric(pivot, fieldName, start, endValue, intervalSize, batch.WorkbookPath, batch.Logger));
}



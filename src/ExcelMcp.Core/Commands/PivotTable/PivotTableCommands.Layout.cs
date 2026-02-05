using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable layout operations - compact, outline, and tabular forms.
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Sets the row layout form for a PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="layoutType">Layout form: 0=Compact, 1=Tabular, 2=Outline</param>
    /// <returns>Result indicating success or failure</returns>
    /// <remarks>
    /// LAYOUT FORMS:
    /// - Compact (0): All row fields in single column with indentation (Excel default)
    /// - Tabular (1): Each field in separate column, subtotals at bottom
    /// - Outline (2): Each field in separate column, subtotals at top
    /// 
    /// SUPPORT:
    /// - Regular PivotTables: Full support for all three forms
    /// - OLAP PivotTables: Full support for all three forms
    /// </remarks>
    public OperationResult SetLayout(IExcelBatch batch, string pivotTableName, int layoutType)
        => ExecuteWithStrategy<OperationResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.SetLayout(pivot, layoutType, batch.WorkbookPath, batch.Logger));
}



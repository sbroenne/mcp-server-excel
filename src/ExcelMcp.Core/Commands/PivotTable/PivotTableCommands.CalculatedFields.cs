using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Calculated Fields operations for PivotTableCommands.
/// Creates custom fields with formulas for Regular PivotTables.
/// OLAP PivotTables use DAX measures instead (see excel_datamodel tool).
/// </summary>
public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotFieldResult CreateCalculatedField(IExcelBatch batch, string pivotTableName,
        string fieldName, string formula)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Determine strategy (OLAP vs Regular)
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);

                // Delegate to strategy with logger
                return strategy.CreateCalculatedField(pivot, fieldName, formula, batch.WorkbookPath, batch.Logger);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc/>
    public CalculatedFieldListResult ListCalculatedFields(IExcelBatch batch, string pivotTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? calculatedFields = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Check if this is an OLAP PivotTable - they don't support calculated fields
                if (PivotTableHelpers.IsOlapPivotTable(pivot))
                {
                    return new CalculatedFieldListResult
                    {
                        Success = false,
                        ErrorMessage = $"PivotTable '{pivotTableName}' is an OLAP PivotTable. OLAP PivotTables do not support calculated fields. Use list-calculated-members instead for Data Model PivotTables."
                    };
                }

                var result = new CalculatedFieldListResult
                {
                    Success = true
                };

                // Get CalculatedFields collection
                calculatedFields = pivot.CalculatedFields();

                int count = calculatedFields.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic? field = null;
                    try
                    {
                        field = calculatedFields.Item(i);

                        var fieldInfo = new CalculatedFieldInfo
                        {
                            Name = field.Name?.ToString() ?? string.Empty,
                            Formula = field.Formula?.ToString() ?? string.Empty,
                            SourceName = field.SourceName?.ToString()
                        };

                        result.CalculatedFields.Add(fieldInfo);
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref calculatedFields);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc/>
    public OperationResult DeleteCalculatedField(IExcelBatch batch, string pivotTableName, string fieldName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? calculatedFields = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Check if this is an OLAP PivotTable - they don't support calculated fields
                if (PivotTableHelpers.IsOlapPivotTable(pivot))
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"PivotTable '{pivotTableName}' is an OLAP PivotTable. OLAP PivotTables do not support calculated fields. Use delete-calculated-member instead for Data Model PivotTables."
                    };
                }

                calculatedFields = pivot.CalculatedFields();

                // Find the calculated field by name
                bool found = false;
                int count = calculatedFields.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic? checkField = null;
                    try
                    {
                        checkField = calculatedFields.Item(i);
                        string name = checkField.Name?.ToString() ?? string.Empty;
                        if (name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                        {
                            field = checkField;
                            checkField = null; // Transfer ownership
                            found = true;
                            break;
                        }
                    }
                    finally
                    {
                        if (checkField != null)
                        {
                            ComUtilities.Release(ref checkField);
                        }
                    }
                }

                if (!found)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Calculated field '{fieldName}' not found in PivotTable '{pivotTableName}'. Use list-calculated-fields to see available calculated fields."
                    };
                }

                // Delete the field
                field.Delete();

                // Refresh the PivotTable
                pivot.RefreshTable();

                return new OperationResult
                {
                    Success = true
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref calculatedFields);
                ComUtilities.Release(ref pivot);
            }
        });
    }
}

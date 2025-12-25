using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable management commands - main partial class with shared state and helper methods
/// </summary>
public partial class PivotTableCommands : IPivotTableCommands
{
    #region Helper Methods

    /// <summary>
    /// Finds a PivotTable by name in the workbook.
    /// Delegates to CoreLookupHelpers.FindPivotTable for the actual lookup.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="pivotTableName">Name of the PivotTable to find</param>
    /// <returns>The PivotTable object if found</returns>
    /// <exception cref="InvalidOperationException">Thrown if PivotTable is not found</exception>
    private static dynamic FindPivotTable(dynamic workbook, string pivotTableName)
        => CoreLookupHelpers.FindPivotTable(workbook, pivotTableName);

    /// <summary>
    /// Executes a strategy-based operation on a PivotTable.
    /// Centralizes the common pattern: find pivot → get strategy → execute → release.
    /// </summary>
    /// <typeparam name="TResult">The result type returned by the strategy operation</typeparam>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="operation">The strategy operation to execute (receives strategy and pivot)</param>
    /// <returns>The result from the strategy operation</returns>
    private static TResult ExecuteWithStrategy<TResult>(
        IExcelBatch batch,
        string pivotTableName,
        Func<IPivotTableFieldStrategy, dynamic, TResult> operation)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return operation(strategy, pivot);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Detects the data type of a field by sampling its values
    /// </summary>
    private static string DetectFieldDataType(dynamic field)
    {
        dynamic? pivotItems = null;
        try
        {
            pivotItems = field.PivotItems;
            var sampleValues = new List<object?>();

            int sampleCount = Math.Min(10, pivotItems.Count);
            for (int i = 1; i <= sampleCount; i++)
            {
                dynamic? item = null;
                try
                {
                    item = pivotItems.Item(i);
                    var value = item.Value;
                    if (value != null)
                        sampleValues.Add(value);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            // Analyze sample values
            if (sampleValues.Count == 0)
                return "Unknown";

            if (sampleValues.All(v => DateTime.TryParse(v?.ToString(), out _)))
                return "Date";
            if (sampleValues.All(v => double.TryParse(v?.ToString(), out _)))
                return "Number";
            if (sampleValues.All(v => bool.TryParse(v?.ToString(), out _)))
                return "Boolean";

            return "Text";
        }
        catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            // PivotItems access failed - cannot determine data type
            return "Unknown";
        }
        finally
        {
            ComUtilities.Release(ref pivotItems);
        }
    }

    /// <summary>
    /// Converts Excel COM constant to AggregationFunction enum
    /// </summary>
    private static AggregationFunction GetAggregationFunctionFromCom(int comFunction)
    {
        return comFunction switch
        {
            XlConsolidationFunction.xlSum => AggregationFunction.Sum,
            XlConsolidationFunction.xlCount => AggregationFunction.Count,
            XlConsolidationFunction.xlAverage => AggregationFunction.Average,
            XlConsolidationFunction.xlMax => AggregationFunction.Max,
            XlConsolidationFunction.xlMin => AggregationFunction.Min,
            XlConsolidationFunction.xlProduct => AggregationFunction.Product,
            XlConsolidationFunction.xlCountNums => AggregationFunction.CountNumbers,
            XlConsolidationFunction.xlStdDev => AggregationFunction.StdDev,
            XlConsolidationFunction.xlStdDevP => AggregationFunction.StdDevP,
            XlConsolidationFunction.xlVar => AggregationFunction.Var,
            XlConsolidationFunction.xlVarP => AggregationFunction.VarP,
            _ => throw new InvalidOperationException($"Unknown COM aggregation function: {comFunction}")
        };
    }

    /// <summary>
    /// Gets all field names from a PivotTable
    /// </summary>
    private static List<string> GetFieldNames(dynamic pivotTable)
    {
        var fieldNames = new List<string>();
        dynamic? pivotFields = null;
        try
        {
            pivotFields = pivotTable.PivotFields;
            for (int i = 1; i <= pivotFields.Count; i++)
            {
                dynamic? field = null;
                try
                {
                    field = pivotFields.Item(i);
                    fieldNames.Add(field.SourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}");
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref pivotFields);
        }
        return fieldNames;
    }

    /// <summary>
    /// Gets unique values from a field for filtering purposes
    /// </summary>
    private static List<string> GetFieldUniqueValues(dynamic field)
    {
        var values = new List<string>();
        dynamic? pivotItems = null;
        try
        {
            pivotItems = field.PivotItems;
            for (int i = 1; i <= pivotItems.Count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = pivotItems.Item(i);
                    string itemName = item.Name?.ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(itemName))
                        values.Add(itemName);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // PivotItems access failed - return partial list
        }
        finally
        {
            ComUtilities.Release(ref pivotItems);
        }
        return values;
    }

    /// <summary>
    /// Gets a field for manipulation, handling both OLAP and regular PivotTables.
    /// For OLAP PivotTables, accesses via CubeFields and returns the corresponding PivotField.
    /// For regular PivotTables, accesses via PivotFields directly.
    /// </summary>
    /// <param name="pivot">The PivotTable object</param>
    /// <param name="fieldName">Name of the field to retrieve</param>
    /// <param name="isOlap">Output parameter indicating if this is an OLAP PivotTable</param>
    /// <returns>The field object that can be manipulated (PivotField)</returns>
    /// <exception cref="InvalidOperationException">Thrown if field is not found</exception>
    /// <remarks>
    /// Microsoft docs: "In OLAP PivotTables, PivotFields do not exist until the corresponding
    /// CubeField is added to the PivotTable." This method handles both architectures.
    /// </remarks>
    private static dynamic GetFieldForManipulation(dynamic pivot, string fieldName, out bool isOlap)
    {
        isOlap = false;
        dynamic? cubeFields = null;

        try
        {
            // Check if this is an OLAP/Data Model PivotTable
            isOlap = PivotTableHelpers.TryGetCubeFields(pivot, out cubeFields);

            if (isOlap)
            {
                // OLAP PivotTable - access via CubeFields
                // CubeField names are hierarchical like "[TableName].[FieldName]" or "[Measures].[MeasureName]"
                // EXACT MATCH ONLY - no partial matching to avoid disambiguation bugs
                dynamic? cubeField = null;
                try
                {
                    // Exact match only - the LLM knows the exact field names
                    try
                    {
                        cubeField = cubeFields.Item(fieldName);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Field not found by exact name
                        cubeField = null;
                    }

                    if (cubeField == null)
                    {
                        throw new InvalidOperationException($"CubeField '{fieldName}' not found in Data Model PivotTable. Use the exact CubeField name (e.g., '[Measures].[ACR]' or '[TableName].[ColumnName]').");
                    }

                    // Get or create the PivotField from the CubeField
                    // Per Microsoft docs: CubeField.PivotFields returns collection of PivotFields for this CubeField
                    dynamic? pivotFields = cubeField.PivotFields;
                    if (pivotFields == null || pivotFields.Count == 0)
                    {
                        // No PivotField exists yet - field hasn't been added to PivotTable
                        // Call CreatePivotFields() to create the PivotFields collection
                        // Per Microsoft docs: "In OLAP PivotTables, PivotFields do not exist until
                        // the corresponding CubeField is added to the PivotTable. The CreatePivotFields()
                        // method enables users to create all PivotFields of a CubeField."
                        ComUtilities.Release(ref pivotFields);
                        cubeField.CreatePivotFields(); // Create PivotFields before manipulation

                        // Now get the newly created PivotFields collection
                        pivotFields = cubeField.PivotFields;
                        if (pivotFields == null || pivotFields.Count == 0)
                        {
                            // Still no PivotFields - this shouldn't happen after CreatePivotFields()
                            ComUtilities.Release(ref pivotFields);
                            throw new InvalidOperationException($"Failed to create PivotFields for CubeField '{fieldName}'");
                        }
                    }

                    // Release PivotFields collection - we don't need it
                    ComUtilities.Release(ref pivotFields);

                    // CRITICAL: Return the CubeField, not the PivotField!
                    // For OLAP, we must set Orientation on the CubeField, not on the PivotField.
                    // Microsoft docs: "CubeField.Orientation returns or sets... the location of the field"
                    // Setting PivotField.Orientation fails with "Unable to set the Orientation property"
                    // DON'T release cubeField - caller needs it!
                    return cubeField;
                }
                finally
                {
                    // Only release cubeField if we didn't return it or a child object
                    // Since we return cubeField, we should NOT release it here
                    // if (cubeField != null)
                    //     ComUtilities.Release(ref cubeField);
                }
            }
            else
            {
                // Regular PivotTable - access via PivotFields directly
                dynamic? pivotFields = null;
                try
                {
                    pivotFields = pivot.PivotFields;
                    return pivotFields.Item(fieldName); // COM will throw if not found
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' not found in PivotTable", ex);
                }
                finally
                {
                    ComUtilities.Release(ref pivotFields);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
        }
    }

    #endregion
}

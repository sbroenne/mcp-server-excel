using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable management commands - main partial class with shared state and helper methods
/// </summary>
public partial class PivotTableCommands : IPivotTableCommands
{
    #region Helper Methods

    /// <summary>
    /// Finds a PivotTable by name in the workbook
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="pivotTableName">Name of the PivotTable to find</param>
    /// <returns>The PivotTable object if found</returns>
    /// <exception cref="InvalidOperationException">Thrown if PivotTable is not found</exception>
    private static dynamic FindPivotTable(dynamic workbook, string pivotTableName)
    {
        dynamic? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic? sheet = null;
                dynamic? pivotTables = null;
                try
                {
                    sheet = sheets.Item(i);
                    pivotTables = sheet.PivotTables;

                    for (int j = 1; j <= pivotTables.Count; j++)
                    {
                        dynamic? pivot = null;
                        try
                        {
                            pivot = pivotTables.Item(j);
                            if (pivot.Name == pivotTableName)
                            {
                                // Found it - return without releasing
                                return pivot;
                            }
                        }
                        finally
                        {
                            if (pivot != null && pivot.Name != pivotTableName)
                            {
                                ComUtilities.Release(ref pivot);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivotTables);
                    ComUtilities.Release(ref sheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }

        throw new InvalidOperationException($"PivotTable '{pivotTableName}' not found in workbook");
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
        catch
        {
            return "Unknown";
        }
        finally
        {
            ComUtilities.Release(ref pivotItems);
        }
    }

    /// <summary>
    /// Validates if an aggregation function is appropriate for a data type
    /// </summary>
    private static bool IsValidAggregationForDataType(AggregationFunction function, string dataType)
    {
        return dataType switch
        {
            "Number" => true, // All functions valid for numbers
            "Date" => function is AggregationFunction.Count or AggregationFunction.CountNumbers or
                      AggregationFunction.Max or AggregationFunction.Min,
            "Text" => function == AggregationFunction.Count,
            "Boolean" => function is AggregationFunction.Count or AggregationFunction.Sum,
            _ => function == AggregationFunction.Count
        };
    }

    /// <summary>
    /// Gets the list of valid aggregation functions for a data type
    /// </summary>
    private static List<string> GetValidAggregationsForDataType(string dataType)
    {
        return dataType switch
        {
            "Number" => ["Sum", "Count", "Average", "Max", "Min", "Product", "CountNumbers", "StdDev", "StdDevP", "Var", "VarP"],
            "Date" => ["Count", "CountNumbers", "Max", "Min"],
            "Text" => ["Count"],
            "Boolean" => ["Count", "Sum"],
            _ => ["Count"]
        };
    }

    /// <summary>
    /// Converts AggregationFunction enum to Excel COM constant
    /// </summary>
    private static int GetComAggregationFunction(AggregationFunction function)
    {
        return function switch
        {
            AggregationFunction.Sum => XlConsolidationFunction.xlSum,
            AggregationFunction.Count => XlConsolidationFunction.xlCount,
            AggregationFunction.Average => XlConsolidationFunction.xlAverage,
            AggregationFunction.Max => XlConsolidationFunction.xlMax,
            AggregationFunction.Min => XlConsolidationFunction.xlMin,
            AggregationFunction.Product => XlConsolidationFunction.xlProduct,
            AggregationFunction.CountNumbers => XlConsolidationFunction.xlCountNums,
            AggregationFunction.StdDev => XlConsolidationFunction.xlStdDev,
            AggregationFunction.StdDevP => XlConsolidationFunction.xlStdDevP,
            AggregationFunction.Var => XlConsolidationFunction.xlVar,
            AggregationFunction.VarP => XlConsolidationFunction.xlVarP,
            _ => throw new InvalidOperationException($"Unsupported aggregation function: {function}")
        };
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
    /// Gets the area name for display purposes
    /// </summary>
    private static string GetAreaName(dynamic orientation)
    {
        int orientationValue = Convert.ToInt32(orientation);
        return orientationValue switch
        {
            XlPivotFieldOrientation.xlHidden => "Hidden",
            XlPivotFieldOrientation.xlRowField => "Row",
            XlPivotFieldOrientation.xlColumnField => "Column",
            XlPivotFieldOrientation.xlPageField => "Filter",
            XlPivotFieldOrientation.xlDataField => "Value",
            _ => $"Unknown({orientationValue})"
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
        catch
        {
            // Ignore errors getting items
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
            try
            {
                cubeFields = pivot.CubeFields;
                isOlap = cubeFields != null && cubeFields.Count > 0;
            }
            catch
            {
                isOlap = false;
            }

            if (isOlap)
            {
                // OLAP PivotTable - access via CubeFields
                // CubeField names may be hierarchical like "[TableName].[FieldName]"
                dynamic? cubeField = null;
                try
                {
                    // Try exact match first
                    try
                    {
                        cubeField = cubeFields.Item(fieldName);
                    }
                    catch
                    {
                        // Try partial match for hierarchical names
                        for (int i = 1; i <= cubeFields.Count; i++)
                        {
                            dynamic? cf = null;
                            try
                            {
                                cf = cubeFields.Item(i);
                                string cfName = cf.Name?.ToString() ?? "";
                                if (cfName.Contains(fieldName, StringComparison.OrdinalIgnoreCase))
                                {
                                    cubeField = cf;
                                    cf = null; // Don't release, we're returning it
                                    break;
                                }
                            }
                            finally
                            {
                                if (cf != null)
                                    ComUtilities.Release(ref cf);
                            }
                        }
                    }

                    if (cubeField == null)
                    {
                        throw new InvalidOperationException($"CubeField '{fieldName}' not found in Data Model PivotTable");
                    }

                    // Get or create the PivotField from the CubeField
                    // Per Microsoft docs: CubeField.PivotFields returns collection of PivotFields for this CubeField
                    dynamic? pivotFields = cubeField.PivotFields;
                    if (pivotFields == null || pivotFields.Count == 0)
                    {
                        // No PivotField exists yet - field hasn't been added to PivotTable
                        // We can still return a reference that will work for adding the field
                        // The field operations will create the PivotField when setting orientation
                        ComUtilities.Release(ref pivotFields);
                        return cubeField; // Return CubeField, operations will adapt
                    }

                    // Return the first PivotField (there's typically only one per CubeField)
                    dynamic pivotField = pivotFields.Item(1);
                    ComUtilities.Release(ref pivotFields);
                    return pivotField;
                }
                finally
                {
                    if (cubeField != null)
                        ComUtilities.Release(ref cubeField);
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

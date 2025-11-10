using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Strategy for OLAP (Online Analytical Processing) PivotTable field operations.
/// Uses CubeFields API for Data Model-based PivotTables.
/// 
/// CRITICAL: In OLAP PivotTables, PivotFields do not exist until the corresponding 
/// CubeField is added to the PivotTable. Must call CreatePivotFields() first.
/// Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.cubefield.createpivotfields
/// </summary>
public class OlapPivotTableFieldStrategy : IPivotTableFieldStrategy
{
    /// <inheritdoc/>
    public bool CanHandle(dynamic pivot)
    {
        try
        {
            // OLAP/Data Model PivotTables have CubeFields collection
            // Note: Don't release COM objects here - PivotTable keeps them alive
            dynamic cubeFields = pivot.CubeFields;
            return cubeFields != null && cubeFields.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public dynamic GetFieldForManipulation(dynamic pivot, string fieldName)
    {
        dynamic? cubeFields = null;
        dynamic? cubeField = null;
        try
        {
            cubeFields = pivot.CubeFields;
            // Try exact match first
            try
            {
                cubeField = cubeFields.Item(fieldName);
            }
            catch
            {
                // Try partial match for hierarchical names (e.g., "[Sales].[Region]" matches "Region")
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
                            cf = null; // Don't release, we're transferring ownership
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
                throw new InvalidOperationException($"Field '{fieldName}' not found in OLAP PivotTable");
            }

            // CRITICAL FIX: CreatePivotFields() must be called before manipulating OLAP fields
            // Without this, PivotFields don't exist and operations fail
            // Reference: https://github.com/NetOfficeFw/NetOffice/search?q=CreatePivotFields
            cubeField.CreatePivotFields();

            return cubeField; // Return CubeField, not PivotField
        }
        catch (Exception ex) when (cubeField == null)
        {
            throw new InvalidOperationException($"Field '{fieldName}' not found in OLAP PivotTable", ex);
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
            // Note: Don't release cubeField - we're returning it
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldListResult ListFields(dynamic pivot, string workbookPath)
    {
        var fields = new List<PivotFieldInfo>();
        dynamic? cubeFields = null;

        try
        {
            cubeFields = pivot.CubeFields;
            int fieldCount = cubeFields.Count;

            for (int i = 1; i <= fieldCount; i++)
            {
                dynamic? cubeField = null;
                try
                {
                    cubeField = cubeFields.Item(i);
                    int orientation = Convert.ToInt32(cubeField.Orientation);

                    // Skip hidden fields
                    if (orientation == XlPivotFieldOrientation.xlHidden)
                        continue;

                    var fieldInfo = new PivotFieldInfo
                    {
                        Name = cubeField.Name?.ToString() ?? $"Field{i}",
                        CustomName = cubeField.Caption?.ToString() ?? "",
                        Area = (PivotFieldArea)orientation,
                        DataType = "Cube" // OLAP fields are always Cube type
                    };

                    // OLAP doesn't support AvailableValues like Regular PivotTables
                    // Values come from OLAP dimension hierarchies

                    fields.Add(fieldInfo);
                }
                catch (Exception ex)
                {
                    // Log but continue with other fields
                    Console.WriteLine($"Error reading OLAP field {i}: {ex.Message}");
                }
                finally
                {
                    ComUtilities.Release(ref cubeField);
                }
            }

            return new PivotFieldListResult
            {
                Success = true,
                Fields = fields,
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldListResult
            {
                Success = false,
                ErrorMessage = $"Failed to list OLAP fields: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddRowField(dynamic pivot, string fieldName, int? position, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // Check if field is already placed
            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
            }

            // CRITICAL: Set Orientation on CubeField, NOT on PivotField
            cubeField.Orientation = XlPivotFieldOrientation.xlRowField;
            if (position.HasValue)
            {
                cubeField.Position = (double)position.Value;
            }

            // Refresh and validate
            pivot.RefreshTable();

            if (cubeField.Orientation != XlPivotFieldOrientation.xlRowField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Row area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Row,
                Position = Convert.ToInt32(cubeField.Position),
                DataType = "Cube",
                AvailableValues = new List<string>(), // OLAP doesn't provide unique values like Regular
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to add OLAP row field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddColumnField(dynamic pivot, string fieldName, int? position, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
            }

            cubeField.Orientation = XlPivotFieldOrientation.xlColumnField;
            if (position.HasValue)
            {
                cubeField.Position = (double)position.Value;
            }

            pivot.RefreshTable();

            if (cubeField.Orientation != XlPivotFieldOrientation.xlColumnField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Column area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Column,
                Position = Convert.ToInt32(cubeField.Position),
                DataType = "Cube",
                AvailableValues = new List<string>(),
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to add OLAP column field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddValueField(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string? customName, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            // OLAP limitation: Cannot add value fields via COM API
            throw new InvalidOperationException(
                $"Cannot add value field '{fieldName}' to OLAP PivotTable. " +
                "OLAP measures must be pre-defined in the Excel Data Model. " +
                "To add measures: (1) Open Data Model in Excel, (2) Create or modify measures, (3) Add to PivotTable manually, (4) Refresh the PivotTable. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.cubefield");
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddFilterField(dynamic pivot, string fieldName, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
            }

            cubeField.Orientation = XlPivotFieldOrientation.xlPageField;
            pivot.RefreshTable();

            if (cubeField.Orientation != XlPivotFieldOrientation.xlPageField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Filter area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Filter,
                Position = Convert.ToInt32(cubeField.Position),
                DataType = "Cube",
                AvailableValues = new List<string>(),
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to add OLAP filter field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult RemoveField(dynamic pivot, string fieldName, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation == XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is not currently placed in any area");
            }

            cubeField.Orientation = XlPivotFieldOrientation.xlHidden;
            pivot.RefreshTable();

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                Area = PivotFieldArea.Hidden,
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to remove OLAP field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SetFieldName(dynamic pivot, string fieldName, string customName, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            // OLAP limitation: Cannot set Caption on CubeFields via COM
            throw new InvalidOperationException(
                $"Cannot rename OLAP field '{fieldName}' to '{customName}'. " +
                "Field names in OLAP PivotTables are derived from the Data Model definition. " +
                "To change field names: (1) Open Data Model in Excel, (2) Rename the dimension/hierarchy, (3) Refresh the PivotTable. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.cubefield.caption");
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFunction(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            // OLAP limitation: Cannot change aggregation function via COM
            throw new InvalidOperationException(
                $"Cannot change aggregation function for OLAP measure '{fieldName}'. " +
                "OLAP measures have aggregation pre-defined in the Data Model. " +
                "To change aggregation: (1) Open Data Model in Excel, (2) Modify the measure definition, (3) Refresh the PivotTable. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.cubefield.function");
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFormat(dynamic pivot, string fieldName, string numberFormat, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            // OLAP limitation: Cannot set NumberFormat on DataFields via COM
            throw new InvalidOperationException(
                $"Cannot set number format for OLAP field '{fieldName}'. " +
                "Number formatting in OLAP PivotTables is controlled by the Data Model definition, not via PivotTable properties. " +
                "To change formatting: (1) Open Data Model in Excel, (2) Modify format settings in the Data Model, (3) Refresh the PivotTable. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.datafield.numberformat");
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldFilterResult SetFieldFilter(dynamic pivot, string fieldName, List<string> filterValues, string workbookPath)
    {
        dynamic? cubeField = null;
        dynamic? pivotFields = null;
        dynamic? pivotField = null;
        dynamic? pivotItems = null;
        try
        {
            // OLAP limitation: Cannot set Visible property on OLAP PivotItems
            throw new InvalidOperationException(
                $"Cannot filter OLAP field '{fieldName}' via PivotItem.Visible property. " +
                "OLAP PivotItems do not support the Visible property. " +
                "To filter OLAP data: (1) Use PivotTable's built-in filter buttons in Excel, (2) Use OLAP Slicers for interactive filtering, or (3) Modify the source Data Model. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.pivotitem.visible");
        }
        catch (Exception ex)
        {
            return new PivotFieldFilterResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotItems);
            ComUtilities.Release(ref pivotField);
            ComUtilities.Release(ref pivotFields);
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SortField(dynamic pivot, string fieldName, SortDirection direction, string workbookPath)
    {
        dynamic? cubeField = null;
        dynamic? pivotFields = null;
        dynamic? pivotField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // OLAP sorting works through PivotField, not CubeField
            pivotFields = cubeField.PivotFields;
            if (pivotFields == null || pivotFields.Count == 0)
            {
                throw new InvalidOperationException($"Cannot sort OLAP field '{fieldName}' - PivotFields not available");
            }

            pivotField = pivotFields.Item(1);

            int sortOrder = direction == SortDirection.Ascending
                ? XlSortOrder.xlAscending
                : XlSortOrder.xlDescending;

            pivotField.AutoSort(sortOrder, fieldName);
            pivot.RefreshTable();

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = (PivotFieldArea)cubeField.Orientation,
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to sort OLAP field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotField);
            ComUtilities.Release(ref pivotFields);
            ComUtilities.Release(ref cubeField);
        }
    }

    #region Helper Methods

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

    #endregion
}

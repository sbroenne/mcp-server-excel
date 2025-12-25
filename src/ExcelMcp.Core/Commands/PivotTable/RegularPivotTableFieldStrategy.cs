using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Strategy for Regular (non-OLAP) PivotTable field operations.
/// Uses PivotFields API for range-based and table-based PivotTables.
/// </summary>
public class RegularPivotTableFieldStrategy : IPivotTableFieldStrategy
{
    /// <inheritdoc/>
    public bool CanHandle(dynamic pivot)
    {
        // Regular PivotTables are NOT OLAP and have PivotFields
        if (PivotTableHelpers.IsOlapPivotTable(pivot))
            return false; // This is OLAP, not regular

        try
        {
            dynamic pivotFields = pivot.PivotFields;
            return pivotFields != null;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public dynamic GetFieldForManipulation(dynamic pivot, string fieldName)
    {
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

    /// <inheritdoc/>
    public PivotFieldListResult ListFields(dynamic pivot, string workbookPath)
    {
        var fields = new List<PivotFieldInfo>();
        dynamic? pivotFields = null;

        try
        {
            pivotFields = pivot.PivotFields;
            int fieldCount = pivotFields.Count;

            for (int i = 1; i <= fieldCount; i++)
            {
                dynamic? field = null;
                try
                {
                    field = pivotFields.Item(i);
                    int orientation = Convert.ToInt32(field.Orientation);

                    var fieldInfo = new PivotFieldInfo
                    {
                        Name = field.SourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}",
                        CustomName = field.Caption?.ToString() ?? "",
                        Area = (PivotFieldArea)orientation,
                        DataType = PivotTableHelpers.DetectFieldDataType(field)
                    };

                    // For value fields, get function from DataFields
                    if (orientation == XlPivotFieldOrientation.xlDataField)
                    {
                        int comFunction = Convert.ToInt32(field.Function);
                        fieldInfo.Function = PivotTableHelpers.GetAggregationFunctionFromCom(comFunction);
                    }

                    fields.Add(fieldInfo);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Individual field access failed - skip this field and continue
                    // This can happen with calculated fields or special field types
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }

            return new PivotFieldListResult
            {
                Success = true,
                Fields = fields,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotFields);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult AddRowField(dynamic pivot, string fieldName, int? position, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            // Check if field is already placed
            int currentOrientation = Convert.ToInt32(field.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            // Add to Row area
            field.Orientation = XlPivotFieldOrientation.xlRowField;
            if (position.HasValue)
            {
                field.Position = (double)position.Value;
            }

            // Refresh and validate
            pivot.RefreshTable();

            if (field.Orientation != XlPivotFieldOrientation.xlRowField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Row area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Row,
                Position = Convert.ToInt32(field.Position),
                DataType = PivotTableHelpers.DetectFieldDataType(field),
                AvailableValues = GetFieldUniqueValues(field),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult AddColumnField(dynamic pivot, string fieldName, int? position, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(field.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            field.Orientation = XlPivotFieldOrientation.xlColumnField;
            if (position.HasValue)
            {
                field.Position = (double)position.Value;
            }

            pivot.RefreshTable();

            if (field.Orientation != XlPivotFieldOrientation.xlColumnField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Column area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Column,
                Position = Convert.ToInt32(field.Position),
                DataType = PivotTableHelpers.DetectFieldDataType(field),
                AvailableValues = GetFieldUniqueValues(field),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult AddValueField(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string? customName, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            // Validate aggregation function for field data type
            string dataType = PivotTableHelpers.DetectFieldDataType(field);
            if (!IsValidAggregationForDataType(aggregationFunction, dataType))
            {
                var validFunctions = GetValidAggregationsForDataType(dataType);
                throw new InvalidOperationException($"Aggregation function '{aggregationFunction}' is not valid for {dataType} field '{fieldName}'. Valid functions: {string.Join(", ", validFunctions)}");
            }

            int currentOrientation = Convert.ToInt32(field.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            field.Orientation = XlPivotFieldOrientation.xlDataField;
            int comFunction = GetComAggregationFunction(aggregationFunction);
            field.Function = comFunction;

            if (!string.IsNullOrEmpty(customName))
            {
                field.Caption = customName;
            }

            pivot.RefreshTable();

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Value,
                Function = aggregationFunction,
                DataType = dataType,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult AddFilterField(dynamic pivot, string fieldName, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(field.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            field.Orientation = XlPivotFieldOrientation.xlPageField;
            pivot.RefreshTable();

            if (field.Orientation != XlPivotFieldOrientation.xlPageField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Filter area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Filter,
                Position = Convert.ToInt32(field.Position),
                DataType = PivotTableHelpers.DetectFieldDataType(field),
                AvailableValues = GetFieldUniqueValues(field),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult RemoveField(dynamic pivot, string fieldName, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(field.Orientation);
            if (currentOrientation == XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is not currently placed in any area");
            }

            field.Orientation = XlPivotFieldOrientation.xlHidden;
            pivot.RefreshTable();

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                Area = PivotFieldArea.Hidden,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult SetFieldName(dynamic pivot, string fieldName, string customName, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);
            field.Caption = customName;

            // NOTE: No RefreshTable() needed - Caption is a visual-only property

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = customName,
                Area = (PivotFieldArea)field.Orientation,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFunction(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            // Find field in DataFields collection (value fields)
            bool foundInDataFields = false;
            for (int i = 1; i <= pivot.DataFields.Count; i++)
            {
                dynamic? dataField = null;
                try
                {
                    dataField = pivot.DataFields.Item(i);
                    string sourceName = dataField.SourceName?.ToString() ?? "";
                    if (sourceName == fieldName)
                    {
                        field = dataField;
                        foundInDataFields = true;
                        break;
                    }
                }
                finally
                {
                    if (!foundInDataFields && dataField != null)
                        ComUtilities.Release(ref dataField);
                }
            }

            if (!foundInDataFields)
            {
                field = GetFieldForManipulation(pivot, fieldName);
                int orientation = Convert.ToInt32(field.Orientation);
                if (orientation != XlPivotFieldOrientation.xlDataField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is not in the Values area. It is in {PivotTableHelpers.GetAreaName(orientation)} area.");
                }
            }

            // Get source field for data type detection
            dynamic? sourceField = GetFieldForManipulation(pivot, fieldName);
            string dataType = PivotTableHelpers.DetectFieldDataType(sourceField);
            ComUtilities.Release(ref sourceField);

            if (!IsValidAggregationForDataType(aggregationFunction, dataType))
            {
                var validFunctions = GetValidAggregationsForDataType(dataType);
                throw new InvalidOperationException($"Aggregation function '{aggregationFunction}' is not valid for {dataType} field '{fieldName}'. Valid functions: {string.Join(", ", validFunctions)}");
            }

            int comFunction = GetComAggregationFunction(aggregationFunction);
            field.Function = comFunction;
            // NOTE: No RefreshTable() needed - Function change persists without refresh
            // (Verified by diagnostic test: FunctionChange_WithoutRefresh_VerifyPersistence)

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Value,
                Function = aggregationFunction,
                DataType = dataType,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFormat(dynamic pivot, string fieldName, string numberFormat, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            // Find field in DataFields collection
            bool foundInDataFields = false;
            for (int i = 1; i <= pivot.DataFields.Count; i++)
            {
                dynamic? dataField = null;
                try
                {
                    dataField = pivot.DataFields.Item(i);
                    string sourceName = dataField.SourceName?.ToString() ?? "";
                    if (sourceName == fieldName)
                    {
                        field = dataField;
                        foundInDataFields = true;
                        break;
                    }
                }
                finally
                {
                    if (!foundInDataFields && dataField != null)
                        ComUtilities.Release(ref dataField);
                }
            }

            if (!foundInDataFields)
            {
                field = GetFieldForManipulation(pivot, fieldName);
                int orientation = Convert.ToInt32(field.Orientation);
                if (orientation != XlPivotFieldOrientation.xlDataField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is not in the Values area. Only value fields can have number formats.");
                }
            }

            field.NumberFormat = numberFormat;

            // NOTE: No RefreshTable() needed - NumberFormat is a visual-only property

            // Read back the format to verify it was set
            string? appliedFormat = null;
            try
            {
                appliedFormat = field.NumberFormat?.ToString();
            }
            catch
            {
                // If we can't read it back, use what we set
                appliedFormat = numberFormat;
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Value,
                NumberFormat = appliedFormat,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldFilterResult SetFieldFilter(dynamic pivot, string fieldName, List<string> filterValues, string workbookPath)
    {
        dynamic? field = null;
        dynamic? pivotItems = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            // Clear all existing filters first
            field.ClearAllFilters();

            // Set visibility based on filter values
            pivotItems = field.PivotItems;
            var availableItems = new List<string>();

            for (int i = 1; i <= pivotItems.Count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = pivotItems.Item(i);
                    string itemName = item.Name?.ToString() ?? "";
                    availableItems.Add(itemName);
                    item.Visible = filterValues.Contains(itemName);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            // NOTE: No RefreshTable() needed - Filter changes persist without refresh
            // (Verified by diagnostic test: Filter_WithoutRefresh_VerifyPersistence)

            return new PivotFieldFilterResult
            {
                Success = true,
                FieldName = fieldName,
                SelectedItems = filterValues,
                AvailableItems = availableItems,
                ShowAll = false,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotItems);
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult SortField(dynamic pivot, string fieldName, SortDirection direction, string workbookPath)
    {
        dynamic? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            int sortOrder = direction == SortDirection.Ascending
                ? XlSortOrder.xlAscending
                : XlSortOrder.xlDescending;

            field.AutoSort(sortOrder, fieldName);

            // NOTE: No RefreshTable() needed - Sorting is a visual-only operation

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = (PivotFieldArea)field.Orientation,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    #region Helper Methods

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
            // Ignore errors
        }
        finally
        {
            ComUtilities.Release(ref pivotItems);
        }
        return values;
    }

    private static bool IsValidAggregationForDataType(AggregationFunction function, string dataType)
    {
        return dataType switch
        {
            "Number" => true,
            "Date" => function is AggregationFunction.Count or AggregationFunction.CountNumbers or
                      AggregationFunction.Max or AggregationFunction.Min,
            "Text" => function == AggregationFunction.Count,
            "Boolean" => function is AggregationFunction.Count or AggregationFunction.Sum,
            _ => function == AggregationFunction.Count
        };
    }

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

    #endregion

    /// <inheritdoc/>
    /// <remarks>
    /// CRITICAL REQUIREMENT: Source data MUST be formatted with date NumberFormat BEFORE creating the PivotTable.
    /// Without proper date formatting, Excel stores dates as serial numbers (e.g., 45672) with "Standard" format,
    /// which prevents date grouping from working.
    ///
    /// Example:
    /// <code>
    /// // Apply date format to source data BEFORE creating PivotTable
    /// sheet.Range["D2:D6"].NumberFormat = "m/d/yyyy";
    /// </code>
    ///
    /// This method groups date fields by Days, Months, Quarters, or Years. Excel automatically creates
    /// hierarchical groupings (e.g., Months + Years) for proper time-based analysis.
    /// </remarks>
    public PivotFieldResult GroupByDate(dynamic pivot, string fieldName, DateGroupingInterval interval, string workbookPath, ILogger? logger = null)
    {
        dynamic? field = null;
        dynamic? singleCell = null;
        try
        {
            // CRITICAL: Refresh PivotTable FIRST to populate field with actual date values
            // Excel needs populated items before grouping can work
            pivot.RefreshTable();

            field = GetFieldForManipulation(pivot, fieldName);

            // CRITICAL: Microsoft documentation states:
            // "The Range object must be a single cell in the PivotTable field's data range"
            // This means a cell from the actual PivotTable BODY (items in the field),
            // NOT the field button area.
            //
            // Source: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range.group?view=excel-pia
            //
            // PivotField.DataRange returns:
            // - For Row/Column/Page fields: "Items in the field" (what we need!)
            // - For Data fields: "Data contained in the field"
            //
            // Use the first cell from field.DataRange - this is where the actual date values appear

            // Get first cell from field.DataRange (items in the field)
            singleCell = field.DataRange.Cells[1, 1];

            // CRITICAL: Periods is a boolean array with 7 elements (Seconds, Minutes, Hours, Days, Months, Quarters, Years)
            // See: https://learn.microsoft.com/en-us/office/vba/api/excel.range.group
            // Element indexes: 1=Seconds, 2=Minutes, 3=Hours, 4=Days, 5=Months, 6=Quarters, 7=Years
            // Excel uses 1-based indexing, C# arrays are 0-based, so index 3 = element 4 = Days
            var periods = new object[] { false, false, false, false, false, false, false };

            switch (interval)
            {
                case DateGroupingInterval.Days:
                    periods[3] = true;      // Element 4 (index 3) = Days
                    break;
                case DateGroupingInterval.Months:
                    periods[4] = true;      // Element 5 (index 4) = Months
                    periods[6] = true;      // Element 7 (index 6) = Years (required for month grouping)
                    break;
                case DateGroupingInterval.Quarters:
                    periods[5] = true;      // Element 6 (index 5) = Quarters
                    periods[6] = true;      // Element 7 (index 6) = Years (required for quarter grouping)
                    break;
                case DateGroupingInterval.Years:
                    periods[6] = true;      // Element 7 (index 6) = Years
                    break;
                default:
                    throw new ArgumentException($"Unknown grouping interval: {interval}");
            }

            // Call Group on single cell, not entire range
            // VBA examples use Start:=True and End:=True to use auto-detected min/max date range
            singleCell.Group(
                Start: true,
                End: true,
                By: Type.Missing,
                Periods: periods
            );

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = (PivotFieldArea)field.Orientation,
                FilePath = workbookPath,
                WorkflowHint = $"Field '{fieldName}' grouped by {interval}. Excel created automatic date hierarchy."
            };
        }
        catch (Exception ex)
        {
            if (logger is not null && logger.IsEnabled(LogLevel.Error))
            {
#pragma warning disable CA1848 // Keep error logging for diagnostics
                logger.LogError(ex, "GroupByDate failed for field '{FieldName}'", fieldName);
#pragma warning restore CA1848
            }
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to group field by date: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref singleCell);
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult GroupByNumeric(dynamic pivot, string fieldName, double? start, double? endValue, double intervalSize, string workbookPath, ILogger? logger = null)
    {
        dynamic? field = null;
        dynamic? singleCell = null;
        try
        {
            // CRITICAL: Refresh PivotTable FIRST to populate field with actual numeric values
            // Excel needs populated items before grouping can work (same as date grouping)
            pivot.RefreshTable();

            field = GetFieldForManipulation(pivot, fieldName);

            // CRITICAL: Microsoft documentation states:
            // "The Range object must be a single cell in the PivotTable field's data range"
            // Same requirement as date grouping - use first cell from field.DataRange
            //
            // Source: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range.group?view=excel-pia
            //
            // For numeric grouping:
            // - By parameter specifies interval size (e.g., 10 for groups of 10)
            // - Start/End parameters define range (null = use field min/max)
            // - Periods parameter is IGNORED (only used for date grouping)

            // Get first cell from field.DataRange (items in the field)
            singleCell = field.DataRange.Cells[1, 1];

            // Convert nullable to object
            // If start/end are null, use true to let Excel auto-detect min/max
            object startValue = start.HasValue ? (object)start.Value : true;
            object endValueObj = endValue.HasValue ? (object)endValue.Value : true;

            // Call Group on single cell
            // For numeric fields, By specifies the interval size
            // Periods is ignored (only used for date grouping)
            singleCell.Group(
                Start: startValue,
                End: endValueObj,
                By: intervalSize,
                Periods: Type.Missing
            );

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = (PivotFieldArea)field.Orientation,
                FilePath = workbookPath,
                WorkflowHint = $"Field '{fieldName}' grouped by intervals of {intervalSize}. Excel created numeric range groups."
            };
        }
        catch (Exception ex)
        {
            if (logger is not null && logger.IsEnabled(LogLevel.Error))
            {
#pragma warning disable CA1848 // Keep error logging for diagnostics
                logger.LogError(ex, "GroupByNumeric failed for field '{FieldName}'", fieldName);
#pragma warning restore CA1848
            }
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to group field numerically: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref singleCell);
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult CreateCalculatedField(dynamic pivot, string fieldName, string formula, string workbookPath, ILogger? logger = null)
    {
        dynamic? calculatedFields = null;
        dynamic? newField = null;

        try
        {
            // CRITICAL: Refresh PivotTable FIRST to ensure field collection is current
            pivot.RefreshTable();

            // Access CalculatedFields collection
            // For regular PivotTables, this collection allows creating custom fields with formulas
            // Formula syntax: Use field names directly (e.g., "=Revenue-Cost")
            // Excel auto-converts field references to proper format
            calculatedFields = pivot.CalculatedFields();

            // Add calculated field with formula
            // UseStandardFormula = true ensures field names are interpreted in US English format
            // regardless of user's locale settings
            newField = calculatedFields.Add(fieldName, formula, true);

            // Refresh again to populate the new calculated field
            pivot.RefreshTable();

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = newField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Hidden, // Calculated fields start hidden until added to values
                Formula = formula,
                FilePath = workbookPath,
                WorkflowHint = $"Calculated field '{fieldName}' created with formula: {formula}. " +
                              "Add to Values area with AddValueField to see results in PivotTable."
            };
        }
        catch (Exception ex)
        {
            if (logger is not null && logger.IsEnabled(LogLevel.Error))
            {
#pragma warning disable CA1848 // Keep error logging for diagnostics
                logger.LogError(ex, "CreateCalculatedField failed for field '{FieldName}' with formula '{Formula}'", fieldName, formula);
#pragma warning restore CA1848
            }
            return new PivotFieldResult
            {
                Success = false,
                FieldName = fieldName,
                Formula = formula,
                ErrorMessage = $"Failed to create calculated field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref newField);
            ComUtilities.Release(ref calculatedFields);
        }
    }

#pragma warning disable CA1848 // Keep logging for diagnostics
    /// <inheritdoc/>
    public OperationResult SetLayout(dynamic pivot, int layoutType, string workbookPath, ILogger? logger = null)
    {
        // xlCompactRow=0, xlTabularRow=1, xlOutlineRow=2
        pivot.RowAxisLayout(layoutType);

        // NOTE: No RefreshTable() needed - Layout is a visual-only property

        if (logger is not null && logger.IsEnabled(LogLevel.Information))
        {
            logger.LogInformation("Set PivotTable layout to {LayoutType}", layoutType);
        }

        return new OperationResult
        {
            Success = true,
            FilePath = workbookPath
        };
    }
#pragma warning restore CA1848

#pragma warning disable CA1848 // Keep logging for diagnostics
    /// <inheritdoc/>
    public PivotFieldResult SetSubtotals(
        dynamic pivot,
        string fieldName,
        bool showSubtotals,
        string workbookPath,
        ILogger? logger = null)
    {
        dynamic? field = null;
        try
        {
            // Get the field from row fields
            dynamic pivotFields = pivot.PivotFields;
            field = pivotFields.Item(fieldName);

            // Set subtotals: index 1 = Automatic
            // If showSubtotals=true, enable Automatic (which sets all others to false)
            // If showSubtotals=false, disable all subtotals
            field.Subtotals[1] = showSubtotals;

            // NOTE: No RefreshTable() needed - Subtotals is a visual-only property

            if (logger is not null && logger.IsEnabled(LogLevel.Information))
            {
                logger.LogInformation("Set subtotals for field {FieldName} to {ShowSubtotals}", fieldName, showSubtotals);
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                FilePath = workbookPath,
                WorkflowHint = showSubtotals
                    ? "Subtotals enabled for field. Automatic function selected based on data type."
                    : "Subtotals disabled for field. Only detail rows visible."
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }
#pragma warning restore CA1848

#pragma warning disable CA1848
    /// <inheritdoc/>
    public OperationResult SetGrandTotals(dynamic pivot, bool showRowGrandTotals, bool showColumnGrandTotals, string workbookPath, ILogger? logger = null)
    {
        try
        {
            // Set row and column grand totals using COM properties
            pivot.RowGrand = showRowGrandTotals;
            pivot.ColumnGrand = showColumnGrandTotals;

            // NOTE: No RefreshTable() needed - GrandTotals are visual-only properties

            if (logger is not null && logger.IsEnabled(LogLevel.Information))
            {
                logger.LogInformation("Set grand totals: Row={RowGrand}, Column={ColumnGrand}", showRowGrandTotals, showColumnGrandTotals);
            }

            return new OperationResult
            {
                Success = true,
                FilePath = workbookPath
            };
        }
        finally
        {
            // No COM objects to release in this method
        }
    }
#pragma warning restore CA1848
}

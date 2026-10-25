using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Strategy for Regular (non-OLAP) PivotTable field operations.
/// Uses PivotFields API for range-based and table-based PivotTables.
/// </summary>
public class RegularPivotTableFieldStrategy : IPivotTableFieldStrategy
{
    private static Excel.PivotField GetFieldForManipulation(Excel.PivotTable pivot, string fieldName)
    {
        Excel.PivotFields? pivotFields = null;
        try
        {
            pivotFields = (Excel.PivotFields)pivot.PivotFields();
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
    public PivotFieldListResult ListFields(Excel.PivotTable pivot, string workbookPath)
    {
        var fields = new List<PivotFieldInfo>();
        Excel.PivotFields? pivotFields = null;

        try
        {
            pivotFields = (Excel.PivotFields)pivot.PivotFields();
            int fieldCount = pivotFields.Count;

            for (int i = 1; i <= fieldCount; i++)
            {
                Excel.PivotField? field = null;
                try
                {
                    field = pivotFields.Item(i);
                    object? sourceName = field.SourceName;
                    if (PivotTableHelpers.IsValuesPlaceholderField(field, sourceName))
                        continue; // Excel's synthetic Values field has an error VARIANT, not a source column.
                    int orientation = (int)field.Orientation;

                    var fieldInfo = new PivotFieldInfo
                    {
                        Name = sourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}",
                        CustomName = field.Caption?.ToString() ?? "",
                        Area = (PivotFieldArea)orientation,
                        DataType = PivotTableHelpers.DetectFieldDataType(field)
                    };

                    // For value fields, get function from DataFields
                    if (orientation == XlPivotFieldOrientation.xlDataField)
                    {
                        int comFunction = (int)field.Function;
                        fieldInfo.Function = PivotTableHelpers.GetAggregationFunctionFromCom(comFunction);
                    }

                    fields.Add(fieldInfo);
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
    public PivotFieldResult AddRowField(Excel.PivotTable pivot, string fieldName, int? position, string workbookPath)
    {
        Excel.PivotField? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            // Check if field is already placed
            int currentOrientation = (int)field.Orientation;
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            // Add to Row area
            field.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            if (position.HasValue)
            {
                field.Position = (double)position.Value;
            }

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // Removed for consistency with OLAP strategy and improved performance (issue #426)

            if (field.Orientation != Excel.XlPivotFieldOrientation.xlRowField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Row area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Row,
                Position = Convert.ToInt32((object)field.Position, System.Globalization.CultureInfo.InvariantCulture),
                DataType = PivotTableHelpers.DetectFieldDataType(field),
                AvailableValues = PivotTableHelpers.GetFieldUniqueValues(field),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult AddColumnField(Excel.PivotTable pivot, string fieldName, int? position, string workbookPath)
    {
        Excel.PivotField? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = (int)field.Orientation;
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            field.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            if (position.HasValue)
            {
                field.Position = (double)position.Value;
            }

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // Removed for consistency with OLAP strategy and improved performance (issue #426)

            if (field.Orientation != Excel.XlPivotFieldOrientation.xlColumnField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Column area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Column,
                Position = Convert.ToInt32((object)field.Position, System.Globalization.CultureInfo.InvariantCulture),
                DataType = PivotTableHelpers.DetectFieldDataType(field),
                AvailableValues = PivotTableHelpers.GetFieldUniqueValues(field),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult AddValueField(Excel.PivotTable pivot, string fieldName, AggregationFunction aggregationFunction, string? customName, string workbookPath)
    {
        Excel.PivotField? field = null;
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

            int currentOrientation = (int)field.Orientation;
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            field.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
            int comFunction = GetComAggregationFunction(aggregationFunction);
            field.Function = (Excel.XlConsolidationFunction)comFunction;

            if (!string.IsNullOrEmpty(customName))
            {
                field.Caption = customName;
            }

            // NOTE: No RefreshTable() needed - field changes take effect immediately
            // Removed for consistency with OLAP strategy and improved performance (issue #426)

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
    public PivotFieldResult AddFilterField(Excel.PivotTable pivot, string fieldName, string workbookPath)
    {
        Excel.PivotField? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = (int)field.Orientation;
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            field.Orientation = Excel.XlPivotFieldOrientation.xlPageField;

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // Removed for consistency with OLAP strategy and improved performance (issue #426)

            if (field.Orientation != Excel.XlPivotFieldOrientation.xlPageField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Filter area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Filter,
                Position = Convert.ToInt32((object)field.Position, System.Globalization.CultureInfo.InvariantCulture),
                DataType = PivotTableHelpers.DetectFieldDataType(field),
                AvailableValues = PivotTableHelpers.GetFieldUniqueValues(field),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult RemoveField(Excel.PivotTable pivot, string fieldName, string workbookPath)
    {
        Excel.PivotField? field = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = (int)field.Orientation;
            if (currentOrientation == XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is not currently placed in any area");
            }

            field.Orientation = Excel.XlPivotFieldOrientation.xlHidden;

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // Removed for consistency with OLAP strategy and improved performance (issue #426)

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
    public PivotFieldResult SetFieldName(Excel.PivotTable pivot, string fieldName, string customName, string workbookPath)
    {
        Excel.PivotField? field = null;
        try
        {
            field = FindDataField((Excel.PivotTable)pivot, fieldName)
                ?? (Excel.PivotField)GetFieldForManipulation(pivot, fieldName);
            field.Caption = customName;

            // NOTE: No RefreshTable() needed - Caption is a visual-only property

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = field.Caption,
                Area = (PivotFieldArea)(int)field.Orientation,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFunction(Excel.PivotTable pivot, string fieldName, AggregationFunction aggregationFunction, string workbookPath)
    {
        Excel.PivotField? field = null;
        Excel.PivotField? sourceField = null;
        try
        {
            field = FindDataField((Excel.PivotTable)pivot, fieldName);
            if (field is null)
            {
                field = GetFieldForManipulation(pivot, fieldName);
                int orientation = (int)field.Orientation;
                if (orientation != XlPivotFieldOrientation.xlDataField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is not in the Values area. It is in {PivotTableHelpers.GetAreaName(orientation)} area.");
                }
            }

            // Get source field for data type detection
            sourceField = GetFieldForManipulation(pivot, fieldName);
            string dataType = PivotTableHelpers.DetectFieldDataType(sourceField);

            if (!IsValidAggregationForDataType(aggregationFunction, dataType))
            {
                var validFunctions = GetValidAggregationsForDataType(dataType);
                throw new InvalidOperationException($"Aggregation function '{aggregationFunction}' is not valid for {dataType} field '{fieldName}'. Valid functions: {string.Join(", ", validFunctions)}");
            }

            int comFunction = GetComAggregationFunction(aggregationFunction);
            field.Function = (Excel.XlConsolidationFunction)comFunction;
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
            ComUtilities.Release(ref sourceField);
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFormat(Excel.PivotTable pivot, string fieldName, string numberFormat, string workbookPath)
    {
        Excel.PivotField? field = null;
        try
        {
            field = FindDataField((Excel.PivotTable)pivot, fieldName);
            if (field is null)
            {
                field = GetFieldForManipulation(pivot, fieldName);
                int orientation = (int)field.Orientation;
                if (orientation != XlPivotFieldOrientation.xlDataField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is not in the Values area. Only value fields can have number formats.");
                }
            }

            ((Excel.PivotField)field!).NumberFormat = numberFormat;

            // NOTE: No RefreshTable() needed - NumberFormat is a visual-only property

            // Read back the format to verify it was set
            string? appliedFormat = ((Excel.PivotField)field!).NumberFormat;

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

    private static Excel.PivotField? FindDataField(Excel.PivotTable pivot, string fieldName)
    {
        Excel.PivotFields? fields = null;
        try
        {
            fields = (Excel.PivotFields)pivot.DataFields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Excel.PivotField? field = null;
                try
                {
                    field = fields.Item(index);
                    if (string.Equals(field.SourceName, fieldName, StringComparison.Ordinal))
                    {
                        var found = field;
                        field = null;
                        return found;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref field);
                }
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref fields);
        }
    }

    /// <inheritdoc/>
    public PivotFieldFilterResult SetFieldFilter(Excel.PivotTable pivot, string fieldName, List<string> filterValues, string workbookPath)
    {
        Excel.PivotField? field = null;
        Excel.PivotItems? pivotItems = null;
        try
        {
            field = GetFieldForManipulation(pivot, fieldName);

            // Clear all existing filters first
            field.ClearAllFilters();

            // Set visibility based on filter values
            pivotItems = (Excel.PivotItems)field.PivotItems();
            var availableItems = new List<string>();

            for (int i = 1; i <= pivotItems.Count; i++)
            {
                Excel.PivotItem? item = null;
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
    public PivotFieldResult SortField(Excel.PivotTable pivot, string fieldName, SortDirection direction, string workbookPath)
    {
        Excel.PivotField? field = null;
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
                Area = (PivotFieldArea)Convert.ToInt32((object)field.Orientation, System.Globalization.CultureInfo.InvariantCulture),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }

    #region Helper Methods

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
        // NOTE: "Unknown" type includes calculated fields where we can't determine the data type.
        // Calculated fields typically produce numeric results, so we allow all numeric aggregations.
        // If the aggregation is truly invalid, Excel COM will throw an error at runtime.
        return dataType switch
        {
            "Number" => ["Sum", "Count", "Average", "Max", "Min", "Product", "CountNumbers", "StdDev", "StdDevP", "Var", "VarP"],
            "Date" => ["Count", "CountNumbers", "Max", "Min"],
            "Text" => ["Count"],
            "Boolean" => ["Count", "Sum"],
            // Unknown = calculated fields or fields we can't inspect - allow all numeric operations
            _ => ["Sum", "Count", "Average", "Max", "Min", "Product", "CountNumbers", "StdDev", "StdDevP", "Var", "VarP"]
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Grouping failure transport returns Success=false with the actual native message; cancellation and disconnection are excluded and propagate.")]
    public PivotFieldResult GroupByDate(Excel.PivotTable pivot, string fieldName, DateGroupingInterval interval, string workbookPath, ILogger? logger = null)
    {
        Excel.PivotField? field = null;
        Excel.Range? dataRange = null;
        Excel.Range? cells = null;
        Excel.Range? singleCell = null;
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
            dataRange = field.DataRange;
            cells = dataRange.Cells;
            singleCell = (Excel.Range)cells[1, 1];

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
                Area = (PivotFieldArea)Convert.ToInt32((object)field.Orientation, System.Globalization.CultureInfo.InvariantCulture),
                FilePath = workbookPath,
                WorkflowHint = $"Field '{fieldName}' grouped by {interval}. Excel created automatic date hierarchy."
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
        {
            if (logger is not null && logger.IsEnabled(LogLevel.Error))
            {
                logger.LogError(ex, "GroupByDate failed for field '{FieldName}'", fieldName);
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
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref dataRange);
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Numeric grouping failure transport returns Success=false with the actual error; cancellation and disconnection are excluded and propagate.")]
    public PivotFieldResult GroupByNumeric(Excel.PivotTable pivot, string fieldName, double? start, double? endValue, double intervalSize, string workbookPath, ILogger? logger = null)
    {
        Excel.PivotField? field = null;
        Excel.Range? dataRange = null;
        Excel.Range? cells = null;
        Excel.Range? singleCell = null;
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
            dataRange = field.DataRange;
            cells = dataRange.Cells;
            singleCell = (Excel.Range)cells[1, 1];

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
                Area = (PivotFieldArea)Convert.ToInt32((object)field.Orientation, System.Globalization.CultureInfo.InvariantCulture),
                FilePath = workbookPath,
                WorkflowHint = $"Field '{fieldName}' grouped by intervals of {intervalSize}. Excel created numeric range groups."
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
        {
            if (logger is not null && logger.IsEnabled(LogLevel.Error))
            {
                logger.LogError(ex, "GroupByNumeric failed for field '{FieldName}'", fieldName);
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
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref dataRange);
            ComUtilities.Release(ref field);
        }
    }

    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Calculated-field failure transport reports Success=false, the requested field/formula and the real error; cancellation and disconnection propagate.")]
    public PivotFieldResult CreateCalculatedField(Excel.PivotTable pivot, string fieldName, string formula, string workbookPath, ILogger? logger = null)
    {
        Excel.CalculatedFields? calculatedFields = null;
        Excel.PivotField? newField = null;

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
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
        {
            if (logger is not null && logger.IsEnabled(LogLevel.Error))
            {
                logger.LogError(ex, "CreateCalculatedField failed for field '{FieldName}' with formula '{Formula}'", fieldName, formula);
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
    /// <inheritdoc/>
    public OperationResult SetLayout(Excel.PivotTable pivot, int rowLayout, string workbookPath, ILogger? logger = null)
    {
        // xlCompactRow=0, xlTabularRow=1, xlOutlineRow=2
        pivot.RowAxisLayout((Excel.XlLayoutRowType)rowLayout);

        // NOTE: No RefreshTable() needed - Layout is a visual-only property

        if (logger is not null && logger.IsEnabled(LogLevel.Information))
        {
            logger.LogInformation("Set PivotTable layout to {LayoutType}", rowLayout);
        }

        return new OperationResult
        {
            Success = true,
            FilePath = workbookPath
        };
    }
    /// <inheritdoc/>
    public PivotFieldResult SetSubtotals(
        Excel.PivotTable pivot,
        string fieldName,
        bool showSubtotals,
        string workbookPath,
        ILogger? logger = null)
    {
        Excel.PivotField? field = null;
        Excel.PivotFields? pivotFields = null;
        try
        {
            // Get the field from row fields
            pivotFields = (Excel.PivotFields)pivot.PivotFields();
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
            ComUtilities.Release(ref pivotFields);
        }
    }
    /// <inheritdoc/>
    public OperationResult SetGrandTotals(Excel.PivotTable pivot, bool showRowGrandTotals, bool showColumnGrandTotals, string workbookPath, ILogger? logger = null)
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
}

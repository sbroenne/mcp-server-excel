using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable field management operations
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Lists all fields in a PivotTable
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldListResult> ListFieldsAsync(IExcelBatch batch, string pivotTableName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? pivotFields = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                pivotFields = pivot.PivotFields;

                var fields = new List<PivotFieldInfo>();

                for (int i = 1; i <= pivotFields.Count; i++)
                {
                    dynamic? field = null;
                    try
                    {
                        field = pivotFields.Item(i);
                        int orientation = field.Orientation;

                        var fieldInfo = new PivotFieldInfo
                        {
                            Name = field.SourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}",
                            CustomName = field.Caption?.ToString() ?? string.Empty,
                            Area = orientation switch
                            {
                                XlPivotFieldOrientation.xlRowField => PivotFieldArea.Row,
                                XlPivotFieldOrientation.xlColumnField => PivotFieldArea.Column,
                                XlPivotFieldOrientation.xlPageField => PivotFieldArea.Filter,
                                XlPivotFieldOrientation.xlDataField => PivotFieldArea.Value,
                                _ => PivotFieldArea.Hidden
                            },
                            Position = orientation != XlPivotFieldOrientation.xlHidden ? Convert.ToInt32(field.Position) : 0,
                            DataType = DetectFieldDataType(field)
                        };

                        // Get function for value fields
                        if (orientation == XlPivotFieldOrientation.xlDataField)
                        {
                            int comFunction = field.Function;
                            fieldInfo.Function = GetAggregationFunctionFromCom(comFunction);
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
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldListResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to list fields: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotFields);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Adds a field to the Row area
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> AddRowFieldAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Validate field exists
                try
                {
                    field = pivot.PivotFields.Item(fieldName);
                }
                catch (Exception)
                {
                    var availableFields = GetFieldNames(pivot);
                    throw new InvalidOperationException($"Field '{fieldName}' not found in PivotTable '{pivotTableName}'. Available fields: {string.Join(", ", availableFields)}");
                }

                // Check if field is already placed
                int currentOrientation = Convert.ToInt32(field.Orientation);
                if (currentOrientation != XlPivotFieldOrientation.xlHidden)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
                }

                // Add to Row area
                field.Orientation = XlPivotFieldOrientation.xlRowField;
                if (position.HasValue)
                {
                    field.Position = (double)position.Value;
                }

                // Refresh and validate placement
                pivot.RefreshTable();

                // Verify field was added successfully
                if (field.Orientation != XlPivotFieldOrientation.xlRowField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Row area. Current orientation: {GetAreaName(field.Orientation)}");
                }

                // Return detailed result
                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = field.Caption?.ToString() ?? fieldName,
                    Area = PivotFieldArea.Row,
                    Position = Convert.ToInt32(field.Position),
                    DataType = DetectFieldDataType(field),
                    AvailableValues = GetFieldUniqueValues(field),
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to add row field: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Adds a field to the Column area
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> AddColumnFieldAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Validate field exists
                try
                {
                    field = pivot.PivotFields.Item(fieldName);
                }
                catch (Exception)
                {
                    var availableFields = GetFieldNames(pivot);
                    throw new InvalidOperationException($"Field '{fieldName}' not found in PivotTable '{pivotTableName}'. Available fields: {string.Join(", ", availableFields)}");
                }

                // Check if field is already placed
                int currentOrientation = field.Orientation;
                if (currentOrientation != XlPivotFieldOrientation.xlHidden)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
                }

                // Add to Column area
                field.Orientation = XlPivotFieldOrientation.xlColumnField;
                if (position.HasValue)
                {
                    field.Position = (double)position.Value;
                }

                // Refresh and validate placement
                pivot.RefreshTable();

                // Verify field was added successfully
                if (field.Orientation != XlPivotFieldOrientation.xlColumnField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Column area. Current orientation: {GetAreaName(field.Orientation)}");
                }

                // Return detailed result
                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = field.Caption?.ToString() ?? fieldName,
                    Area = PivotFieldArea.Column,
                    Position = Convert.ToInt32(field.Position),
                    DataType = DetectFieldDataType(field),
                    AvailableValues = GetFieldUniqueValues(field),
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to add column field: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Adds a field to the Values area with aggregation
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> AddValueFieldAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction function = AggregationFunction.Sum,
        string? customName = null)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Validate field exists
                try
                {
                    field = pivot.PivotFields.Item(fieldName);
                }
                catch (Exception)
                {
                    var availableFields = GetFieldNames(pivot);
                    throw new InvalidOperationException($"Field '{fieldName}' not found in PivotTable '{pivotTableName}'. Available fields: {string.Join(", ", availableFields)}");
                }

                // Validate aggregation function for field data type
                string dataType = DetectFieldDataType(field);
                if (!IsValidAggregationForDataType(function, dataType))
                {
                    var validFunctions = GetValidAggregationsForDataType(dataType);
                    throw new InvalidOperationException($"Aggregation function '{function}' is not valid for {dataType} field '{fieldName}'. Valid functions: {string.Join(", ", validFunctions)}");
                }

                // Add to Values area
                field.Orientation = XlPivotFieldOrientation.xlDataField;

                // Set aggregation function with COM constant
                int comFunction = GetComAggregationFunction(function);
                field.Function = comFunction;

                // Set custom name if provided
                if (!string.IsNullOrEmpty(customName))
                {
                    field.Caption = customName;
                }

                // Refresh and validate
                pivot.RefreshTable();

                // Return detailed result
                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = field.Caption?.ToString() ?? fieldName,
                    Area = PivotFieldArea.Value,
                    Function = function,
                    DataType = dataType,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to add value field: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Adds a field to the Filter area
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> AddFilterFieldAsync(IExcelBatch batch, string pivotTableName,
        string fieldName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Validate field exists
                try
                {
                    field = pivot.PivotFields.Item(fieldName);
                }
                catch (Exception)
                {
                    var availableFields = GetFieldNames(pivot);
                    throw new InvalidOperationException($"Field '{fieldName}' not found in PivotTable '{pivotTableName}'. Available fields: {string.Join(", ", availableFields)}");
                }

                // Check if field is already placed
                int currentOrientation = field.Orientation;
                if (currentOrientation != XlPivotFieldOrientation.xlHidden)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
                }

                // Add to Filter area
                field.Orientation = XlPivotFieldOrientation.xlPageField;

                // Refresh and validate placement
                pivot.RefreshTable();

                // Verify field was added successfully
                if (field.Orientation != XlPivotFieldOrientation.xlPageField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Filter area. Current orientation: {GetAreaName(field.Orientation)}");
                }

                // Return detailed result
                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = field.Caption?.ToString() ?? fieldName,
                    Area = PivotFieldArea.Filter,
                    Position = Convert.ToInt32(field.Position),
                    DataType = DetectFieldDataType(field),
                    AvailableValues = GetFieldUniqueValues(field),
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to add filter field: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Removes a field from any area
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> RemoveFieldAsync(IExcelBatch batch, string pivotTableName,
        string fieldName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Validate field exists
                try
                {
                    field = pivot.PivotFields.Item(fieldName);
                }
                catch (Exception)
                {
                    var availableFields = GetFieldNames(pivot);
                    throw new InvalidOperationException($"Field '{fieldName}' not found in PivotTable '{pivotTableName}'. Available fields: {string.Join(", ", availableFields)}");
                }

                // Check if field is currently placed
                int currentOrientation = field.Orientation;
                if (currentOrientation == XlPivotFieldOrientation.xlHidden)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is not currently placed in any area");
                }

                // Remove from area
                field.Orientation = XlPivotFieldOrientation.xlHidden;

                // Refresh
                pivot.RefreshTable();

                // Return result
                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    Area = PivotFieldArea.Hidden,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to remove field: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sets the aggregation function for a value field
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> SetFieldFunctionAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction function)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                field = pivot.PivotFields.Item(fieldName);

                // Verify field is in Values area
                if (field.Orientation != XlPivotFieldOrientation.xlDataField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is not in the Values area. It is in {GetAreaName(field.Orientation)} area.");
                }

                // Validate function for data type
                string dataType = DetectFieldDataType(field);
                if (!IsValidAggregationForDataType(function, dataType))
                {
                    var validFunctions = GetValidAggregationsForDataType(dataType);
                    throw new InvalidOperationException($"Aggregation function '{function}' is not valid for {dataType} field '{fieldName}'. Valid functions: {string.Join(", ", validFunctions)}");
                }

                // Set function
                int comFunction = GetComAggregationFunction(function);
                field.Function = comFunction;

                // Refresh
                pivot.RefreshTable();

                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = field.Caption?.ToString() ?? fieldName,
                    Area = PivotFieldArea.Value,
                    Function = function,
                    DataType = dataType,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to set field function: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sets custom name for a field
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> SetFieldNameAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, string customName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                field = pivot.PivotFields.Item(fieldName);

                // Set custom name
                field.Caption = customName;

                // Refresh
                pivot.RefreshTable();

                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = customName,
                    Area = (PivotFieldArea)field.Orientation,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to set field name: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sets number format for a value field
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotFieldResult> SetFieldFormatAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, string numberFormat)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                field = pivot.PivotFields.Item(fieldName);

                // Verify field is in Values area
                if (field.Orientation != XlPivotFieldOrientation.xlDataField)
                {
                    throw new InvalidOperationException($"Field '{fieldName}' is not in the Values area. Only value fields can have number formats.");
                }

                // Set number format
                field.NumberFormat = numberFormat;

                // Refresh
                pivot.RefreshTable();

                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = field.Caption?.ToString() ?? fieldName,
                    Area = PivotFieldArea.Value,
                    NumberFormat = numberFormat,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to set field format: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }
}

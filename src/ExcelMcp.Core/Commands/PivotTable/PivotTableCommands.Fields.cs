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
    public PivotFieldListResult ListFields(IExcelBatch batch, string pivotTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? pivotFields = null;
            dynamic? cubeFields = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            // Check if this is an OLAP/Data Model PivotTable
            bool isOlap = false;
            try
            {
                cubeFields = pivot.CubeFields;
                isOlap = cubeFields != null && cubeFields.Count > 0;
            }
            catch
            {
                isOlap = false;
            }

            try
            {
                // For OLAP PivotTables, use CubeFields instead of PivotFields
                if (isOlap)
                {
                    return ListCubeFieldsAsync(cubeFields, batch.WorkbookPath);
                }
                else
                {
                    // Regular PivotTable - use PivotFields
                    pivotFields = pivot.PivotFields;
                    return ListRegularFieldsAsync(pivotFields, batch.WorkbookPath);
                }
            }
            finally
            {
                ComUtilities.Release(ref cubeFields);
                ComUtilities.Release(ref pivotFields);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Lists fields from OLAP/Data Model PivotTable using CubeFields
    /// </summary>
    private static PivotFieldListResult ListCubeFieldsAsync(dynamic cubeFields, string filePath)
    {
        var fields = new List<PivotFieldInfo>();

        try
        {
            int fieldCount = cubeFields.Count;

            for (int i = 1; i <= fieldCount; i++)
            {
                dynamic? cubeField = null;
                try
                {
                    cubeField = cubeFields.Item(i);

                    // Get field name
                    string fieldName;
                    try
                    {
                        fieldName = cubeField.Name?.ToString() ?? $"CubeField{i}";
                    }
                    catch
                    {
                        fieldName = $"CubeField{i}";
                    }

                    // Get orientation - for CubeFields, check if it has a corresponding PivotField
                    int orientation = XlPivotFieldOrientation.xlHidden;
                    try
                    {
                        dynamic? pivotField = cubeField.PivotFields?.Item(1);
                        if (pivotField != null)
                        {
                            orientation = Convert.ToInt32(pivotField.Orientation);
                            ComUtilities.Release(ref pivotField);
                        }
                    }
                    catch
                    {
                        orientation = XlPivotFieldOrientation.xlHidden;
                    }

                    var fieldInfo = new PivotFieldInfo
                    {
                        Name = fieldName,
                        Area = orientation switch
                        {
                            XlPivotFieldOrientation.xlRowField => PivotFieldArea.Row,
                            XlPivotFieldOrientation.xlColumnField => PivotFieldArea.Column,
                            XlPivotFieldOrientation.xlPageField => PivotFieldArea.Filter,
                            XlPivotFieldOrientation.xlDataField => PivotFieldArea.Value,
                            _ => PivotFieldArea.Hidden
                        },
                        CustomName = string.Empty,
                        Position = 0,
                        DataType = "Cube"
                    };

                    fields.Add(fieldInfo);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Warning: Failed to read cube field {i}: {ex.Message}");
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
                FilePath = filePath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldListResult
            {
                Success = false,
                ErrorMessage = $"Failed to list cube fields: {ex.Message}",
                FilePath = filePath
            };
        }
    }

    /// <summary>
    /// Lists fields from regular PivotTable using PivotFields
    /// </summary>
    private static PivotFieldListResult ListRegularFieldsAsync(dynamic pivotFields, string filePath)
    {
        var fields = new List<PivotFieldInfo>();

        try
        {
            int fieldCount;
            try
            {
                fieldCount = pivotFields.Count;
            }
            catch (Exception ex)
            {
                return new PivotFieldListResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to access PivotFields.Count: {ex.Message} (0x{ex.HResult:X})",
                    FilePath = filePath
                };
            }

            for (int i = 1; i <= fieldCount; i++)
            {
                dynamic? field = null;
                try
                {
                    field = pivotFields.Item(i);

                    // Get field name with defensive handling (can throw on Data Model fields)
                    string fieldName;
                    try
                    {
                        fieldName = field.SourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}";
                    }
                    catch
                    {
                        fieldName = $"Field{i}";
                    }

                    // Get orientation with defensive handling
                    int orientation;
                    try
                    {
                        orientation = Convert.ToInt32(field.Orientation);
                    }
                    catch
                    {
                        orientation = XlPivotFieldOrientation.xlHidden;
                    }

                    var fieldInfo = new PivotFieldInfo
                    {
                        Name = fieldName,
                        Area = orientation switch
                        {
                            XlPivotFieldOrientation.xlRowField => PivotFieldArea.Row,
                            XlPivotFieldOrientation.xlColumnField => PivotFieldArea.Column,
                            XlPivotFieldOrientation.xlPageField => PivotFieldArea.Filter,
                            XlPivotFieldOrientation.xlDataField => PivotFieldArea.Value,
                            _ => PivotFieldArea.Hidden
                        }
                    };

                    // CustomName - defensive
                    try
                    {
                        fieldInfo.CustomName = field.Caption?.ToString() ?? string.Empty;
                    }
                    catch
                    {
                        fieldInfo.CustomName = string.Empty;
                    }

                    // Position - defensive
                    try
                    {
                        fieldInfo.Position = orientation != XlPivotFieldOrientation.xlHidden ? Convert.ToInt32(field.Position) : 0;
                    }
                    catch
                    {
                        fieldInfo.Position = 0;
                    }

                    // DataType - defensive
                    try
                    {
                        fieldInfo.DataType = DetectFieldDataType(field);
                    }
                    catch
                    {
                        fieldInfo.DataType = "Unknown";
                    }

                    // Get function for value fields - defensive
                    if (orientation == XlPivotFieldOrientation.xlDataField)
                    {
                        try
                        {
                            int comFunction = Convert.ToInt32(field.Function);
                            fieldInfo.Function = GetAggregationFunctionFromCom(comFunction);
                        }
                        catch
                        {
                            fieldInfo.Function = AggregationFunction.Sum; // Default
                        }
                    }

                    fields.Add(fieldInfo);
                }
                catch (Exception ex)
                {
                    // Log but continue - don't let one bad field break the entire list
                    Console.Error.WriteLine($"Warning: Failed to read field {i}: {ex.Message}");
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
                FilePath = filePath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldListResult
            {
                Success = false,
                ErrorMessage = $"Failed to list regular fields: {ex.Message}",
                FilePath = filePath
            };
        }
    }

    /// <summary>
    /// Adds a field to the Row area
    /// </summary>
    public PivotFieldResult AddRowField(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.AddRowField(pivot, fieldName, position, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Adds a field to the Column area
    /// </summary>
    public PivotFieldResult AddColumnField(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.AddColumnField(pivot, fieldName, position, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Adds a field to the Values area with aggregation
    /// </summary>
    public PivotFieldResult AddValueField(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction aggregationFunction = AggregationFunction.Sum,
        string? customName = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.AddValueField(pivot, fieldName, aggregationFunction, customName, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Adds a field to the Filter area
    /// </summary>
    public PivotFieldResult AddFilterField(IExcelBatch batch, string pivotTableName,
        string fieldName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.AddFilterField(pivot, fieldName, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Removes a field from any area
    /// </summary>
    public PivotFieldResult RemoveField(IExcelBatch batch, string pivotTableName,
        string fieldName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.RemoveField(pivot, fieldName, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sets the aggregation function for a value field
    /// </summary>
    public PivotFieldResult SetFieldFunction(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction aggregationFunction)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SetFieldFunction(pivot, fieldName, aggregationFunction, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sets custom name for a field
    /// </summary>
    public PivotFieldResult SetFieldName(IExcelBatch batch, string pivotTableName,
        string fieldName, string customName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SetFieldName(pivot, fieldName, customName, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sets number format for a value field
    /// </summary>
    public PivotFieldResult SetFieldFormat(IExcelBatch batch, string pivotTableName,
        string fieldName, string numberFormat)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Translate US format codes to locale-specific codes
                var translatedFormat = ctx.FormatTranslator.TranslateToLocale(numberFormat);

                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SetFieldFormat(pivot, fieldName, translatedFormat, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }
}


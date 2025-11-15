using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Read operations - List, View, Export
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public DataModelTableListResult ListTables(IExcelBatch batch)
    {
        var result = new DataModelTableListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            // Check if workbook has Data Model
            if (!HasDataModelTables(ctx.Book))
            {
                // Empty Data Model is valid - return empty list (LLM-friendly)
                result.Success = true;
                result.Tables = [];
                return result;
            }

            dynamic? model = null;
            try
            {
                model = ctx.Book.Model;

                ForEachTable(model, (Action<dynamic, int>)((table, index) =>
                {
                    var tableInfo = new DataModelTableInfo
                    {
                        Name = ComUtilities.SafeGetString(table, "Name"),
                        SourceName = ComUtilities.SafeGetString(table, "SourceName"),
                        RecordCount = ComUtilities.SafeGetInt(table, "RecordCount")
                    };

                    result.Tables.Add(tableInfo);
                }));

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("List tables", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelMeasureListResult ListMeasures(IExcelBatch batch, string? tableName = null)
    {
        var result = new DataModelMeasureListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                // Iterate through all measures (they're at model level)
                ForEachMeasure(model, (Action<dynamic, int>)((measure, index) =>
                {
                    // Get the table name for this measure
                    string measureTableName = string.Empty;
                    dynamic? associatedTable = null;
                    try
                    {
                        associatedTable = measure.AssociatedTable;
                        measureTableName = associatedTable?.Name?.ToString() ?? string.Empty;
                    }
                    finally
                    {
                        ComUtilities.Release(ref associatedTable);
                    }

                    // Skip if filtering by table and this measure isn't in that table
                    if (tableName != null && !measureTableName.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    string formula = ComUtilities.SafeGetString(measure, "Formula");
                    string preview = formula.Length > 80 ? formula[..77] + "..." : formula;

                    var measureInfo = new DataModelMeasureInfo
                    {
                        Name = ComUtilities.SafeGetString(measure, "Name"),
                        Table = measureTableName,
                        FormulaPreview = preview,
                        Description = ComUtilities.SafeGetString(measure, "Description")
                    };

                    result.Measures.Add(measureInfo);
                }));

                // Check if table filter was specified but not found
                if (tableName != null && result.Measures.Count == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(tableName);
                    return result;
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("listing measures", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelMeasureViewResult Read(IExcelBatch batch, string measureName)
    {
        var result = new DataModelMeasureViewResult
        {
            FilePath = batch.WorkbookPath,
            MeasureName = measureName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    var measureNames = GetModelMeasureNames(model);
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.MeasureNotFound(measureName);

                    // Suggest similar measure names
                    var suggestions = new List<string>();
                    foreach (var m in measureNames)
                    {
                        if (m.Contains(measureName, StringComparison.OrdinalIgnoreCase))
                        {
                            suggestions.Add($"Try measure: {m}");
                            if (suggestions.Count >= 3) break;
                        }
                    }

                    return result;
                }

                // Get measure details using safe helpers
                result.DaxFormula = ComUtilities.SafeGetString(measure, "Formula");
                result.Description = ComUtilities.SafeGetString(measure, "Description");
                result.CharacterCount = result.DaxFormula.Length;
                result.TableName = GetMeasureTableName(model, measureName) ?? "";

                // Try to get format information
                try
                {
                    dynamic? formatInfo = measure.FormatInformation;
                    if (formatInfo != null)
                    {
                        try
                        {
                            result.FormatString = formatInfo.FormatString?.ToString();
                        }
                        catch { /* FormatString may not be accessible */ }
                        finally
                        {
                            ComUtilities.Release(ref formatInfo);
                        }
                    }
                }
                catch { /* FormatInformation may not be available in all Excel versions */ }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("viewing measure", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public OperationResult ExportMeasure(IExcelBatch batch, string measureName, string outputFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-export-measure"
        };

        // Validate and normalize output file path
        try
        {
            outputFile = PathValidator.ValidateOutputFile(outputFile, nameof(outputFile), allowOverwrite: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid output file path: {ex.Message}";
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.MeasureNotFound(measureName);
                    return result;
                }

                // Get measure details using safe helpers
                string daxFormula = ComUtilities.SafeGetString(measure, "Formula");
                string description = ComUtilities.SafeGetString(measure, "Description");
                string tableName = GetMeasureTableName(model, measureName) ?? "";
                string? formatString = null;

                // Try to get format information
                try
                {
                    dynamic? formatInfo = measure.FormatInformation;
                    if (formatInfo != null)
                    {
                        try
                        {
                            formatString = formatInfo.FormatString?.ToString();
                        }
                        finally
                        {
                            ComUtilities.Release(ref formatInfo);
                        }
                    }
                }
                catch { }

                // Build DAX file content with metadata
                var daxContent = new System.Text.StringBuilder();
                daxContent.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"-- Measure: {measureName}");
                daxContent.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"-- Table: {tableName}");
                if (!string.IsNullOrEmpty(description))
                {
                    daxContent.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"-- Description: {description}");
                }
                if (!string.IsNullOrEmpty(formatString))
                {
                    daxContent.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"-- Format: {formatString}");
                }
                daxContent.AppendLine();
                daxContent.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"{measureName} :=");
                daxContent.AppendLine(daxFormula);

                // Write to file
                File.WriteAllText(outputFile, daxContent.ToString());

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("exporting measure", ex.Message);
                return result;
            }
            finally
            {
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref model);
            }
        });
    }

    /// <inheritdoc />
    public DataModelRelationshipListResult ListRelationships(IExcelBatch batch)
    {
        var result = new DataModelRelationshipListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                ForEachRelationship(model, (Action<dynamic, int>)((relationship, index) =>
                {
                    var relInfo = new DataModelRelationshipInfo
                    {
                        FromTable = ComUtilities.SafeGetString(relationship.ForeignKeyColumn?.Parent, "Name"),
                        FromColumn = ComUtilities.SafeGetString(relationship.ForeignKeyColumn, "Name"),
                        ToTable = ComUtilities.SafeGetString(relationship.PrimaryKeyColumn?.Parent, "Name"),
                        ToColumn = ComUtilities.SafeGetString(relationship.PrimaryKeyColumn, "Name"),
                        IsActive = relationship.Active ?? false
                    };

                    result.Relationships.Add(relInfo);
                }));

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("listing relationships", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelTableColumnsResult ListColumns(IExcelBatch batch, string tableName)
    {
        var result = new DataModelTableColumnsResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? table = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the table
                table = FindModelTable(model, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(tableName);
                    return result;
                }

                // Iterate through columns
                ComUtilities.ForEachColumn(table, (Action<dynamic, int>)((column, index) =>
                {
                    bool isCalculated = false;
                    try
                    {
                        // IsCalculatedColumn property may not exist in older Excel versions
                        isCalculated = column.IsCalculatedColumn ?? false;
                    }
                    catch (Exception ex) when (ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException
                                            or System.Runtime.InteropServices.COMException)
                    {
                        // Ignore - property not available in this Excel version
                        isCalculated = false;
                    }

                    var columnInfo = new DataModelColumnInfo
                    {
                        Name = ComUtilities.SafeGetString(column, "Name"),
                        DataType = ComUtilities.SafeGetString(column, "DataType"),
                        IsCalculated = isCalculated
                    };

                    result.Columns.Add(columnInfo);
                }));

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed($"listing columns for table '{tableName}'", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelTableViewResult ReadTable(IExcelBatch batch, string tableName)
    {
        var result = new DataModelTableViewResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? table = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the table
                table = FindModelTable(model, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(tableName);
                    return result;
                }

                // Get table properties
                result.SourceName = ComUtilities.SafeGetString(table, "SourceName");
                result.RecordCount = ComUtilities.SafeGetInt(table, "RecordCount");

                // Get columns
                ComUtilities.ForEachColumn(table, (Action<dynamic, int>)((column, index) =>
                {
                    bool isCalculated = false;
                    try
                    {
                        // IsCalculatedColumn property may not exist in older Excel versions
                        isCalculated = column.IsCalculatedColumn ?? false;
                    }
                    catch (Exception ex) when (ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException
                                            or System.Runtime.InteropServices.COMException)
                    {
                        // Ignore - property not available in this Excel version
                        isCalculated = false;
                    }

                    var columnInfo = new DataModelColumnInfo
                    {
                        Name = ComUtilities.SafeGetString(column, "Name"),
                        DataType = ComUtilities.SafeGetString(column, "DataType"),
                        IsCalculated = isCalculated
                    };

                    result.Columns.Add(columnInfo);
                }));

                // Count measures in this table
                result.MeasureCount = 0;
                ForEachMeasure(model, (Action<dynamic, int>)((measure, index) =>
                {
                    string measureTableName = ComUtilities.SafeGetString(measure.AssociatedTable, "Name");
                    if (string.Equals(measureTableName, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        result.MeasureCount++;
                    }
                }));

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed($"viewing table '{tableName}'", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelInfoResult ReadInfo(IExcelBatch batch)
    {
        var result = new DataModelInfoResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                // Count tables and sum rows
                int totalRows = 0;
                ForEachTable(model, (Action<dynamic, int>)((table, index) =>
                {
                    result.TableCount++;
                    totalRows += ComUtilities.SafeGetInt(table, "RecordCount");
                    result.TableNames.Add(ComUtilities.SafeGetString(table, "Name"));
                }));
                result.TotalRows = totalRows;

                // Count measures
                ForEachMeasure(model, (Action<dynamic, int>)((measure, index) =>
                {
                    result.MeasureCount++;
                }));

                // Count relationships
                ForEachRelationship(model, (Action<dynamic, int>)((relationship, index) =>
                {
                    result.RelationshipCount++;
                }));

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("getting model info", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }
}


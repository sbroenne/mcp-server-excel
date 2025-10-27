using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using Sbroenne.ExcelMcp.ComInterop.Session;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model management commands - Core data layer (no console output)
/// Provides read-only access to Excel Data Model (PowerPivot) objects
/// </summary>
public class DataModelCommands : IDataModelCommands
{
    /// <inheritdoc />
    public async Task<DataModelTableListResult> ListTablesAsync(IExcelBatch batch)
    {
        var result = new DataModelTableListResult { FilePath = batch.WorkbookPath };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model. Load data to Data Model first using Power Query or external data sources.";
                    return result;
                }

                model = ctx.Book.Model;
                dynamic? modelTables = null;
                try
                {
                    modelTables = model.ModelTables;
                    int count = modelTables.Count;

                    for (int i = 1; i <= count; i++)
                    {
                        dynamic? table = null;
                        try
                        {
                            table = modelTables.Item(i);

                            var tableInfo = new DataModelTableInfo
                            {
                                Name = table.Name?.ToString() ?? "",
                                SourceName = table.SourceName?.ToString() ?? "",
                                RecordCount = table.RecordCount ?? 0
                            };

                            // Try to get refresh date (may not always be available)
                            try
                            {
                                DateTime? refreshDate = table.RefreshDate;
                                if (refreshDate.HasValue)
                                {
                                    tableInfo = new DataModelTableInfo
                                    {
                                        Name = tableInfo.Name,
                                        SourceName = tableInfo.SourceName,
                                        RecordCount = tableInfo.RecordCount,
                                        RefreshDate = refreshDate
                                    };
                                }
                            }
                            catch { /* RefreshDate not always accessible */ }

                            result.Tables.Add(tableInfo);
                        }
                        finally
                        {
                            ComUtilities.Release(ref table);
                        }
                    }

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error accessing Data Model: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<DataModelMeasureListResult> ListMeasuresAsync(IExcelBatch batch, string? tableName = null)
    {
        var result = new DataModelMeasureListResult { FilePath = batch.WorkbookPath };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = ctx.Book.Model;
                dynamic? modelTables = null;
                try
                {
                    modelTables = model.ModelTables;

                    for (int t = 1; t <= modelTables.Count; t++)
                    {
                        dynamic? table = null;
                        dynamic? measures = null;
                        try
                        {
                            table = modelTables.Item(t);
                            string currentTableName = table.Name?.ToString() ?? "";

                            // Skip if filtering by table and this isn't the table
                            if (tableName != null && !currentTableName.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            measures = table.ModelMeasures;

                            for (int m = 1; m <= measures.Count; m++)
                            {
                                dynamic? measure = null;
                                try
                                {
                                    measure = measures.Item(m);
                                    string formula = measure.Formula?.ToString() ?? "";
                                    string preview = formula.Length > 80 ? formula[..77] + "..." : formula;

                                    var measureInfo = new DataModelMeasureInfo
                                    {
                                        Name = measure.Name?.ToString() ?? "",
                                        Table = currentTableName,
                                        FormulaPreview = preview,
                                        Description = measure.Description?.ToString()
                                    };

                                    result.Measures.Add(measureInfo);
                                }
                                finally
                                {
                                    ComUtilities.Release(ref measure);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref measures);
                            ComUtilities.Release(ref table);
                        }
                    }

                    // Check if table filter was specified but not found
                    if (tableName != null && result.Measures.Count == 0)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Table '{tableName}' not found in Data Model or contains no measures.";
                        return result;
                    }

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing measures: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<DataModelMeasureViewResult> ViewMeasureAsync(IExcelBatch batch, string measureName)
    {
        var result = new DataModelMeasureViewResult
        {
            FilePath = batch.WorkbookPath,
            MeasureName = measureName
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = ComUtilities.FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    var measureNames = DataModelHelpers.GetModelMeasureNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Measure '{measureName}' not found in Data Model.";

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

                    if (suggestions.Any())
                    {
                        result.SuggestedNextActions = suggestions;
                    }

                    return result;
                }

                // Get measure details
                result.DaxFormula = measure.Formula?.ToString() ?? "";
                result.Description = measure.Description?.ToString();
                result.CharacterCount = result.DaxFormula.Length;
                result.TableName = DataModelHelpers.GetMeasureTableName(model, measureName) ?? "";

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
                result.ErrorMessage = $"Error viewing measure: {ex.Message}";
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
    public async Task<OperationResult> ExportMeasureAsync(IExcelBatch batch, string measureName, string outputFile)
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

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = ComUtilities.FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Measure '{measureName}' not found in Data Model.";
                    return result;
                }

                // Get measure details
                string daxFormula = measure.Formula?.ToString() ?? "";
                string? description = measure.Description?.ToString();
                string tableName = DataModelHelpers.GetMeasureTableName(model, measureName) ?? "";
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
                daxContent.AppendLine($"-- Measure: {measureName}");
                daxContent.AppendLine($"-- Table: {tableName}");
                if (!string.IsNullOrEmpty(description))
                {
                    daxContent.AppendLine($"-- Description: {description}");
                }
                if (!string.IsNullOrEmpty(formatString))
                {
                    daxContent.AppendLine($"-- Format: {formatString}");
                }
                daxContent.AppendLine();
                daxContent.AppendLine($"{measureName} :=");
                daxContent.AppendLine(daxFormula);

                // Write to file
                File.WriteAllText(outputFile, daxContent.ToString());

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting measure: {ex.Message}";
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
    public async Task<DataModelRelationshipListResult> ListRelationshipsAsync(IExcelBatch batch)
    {
        var result = new DataModelRelationshipListResult { FilePath = batch.WorkbookPath };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = ctx.Book.Model;
                dynamic? relationships = null;
                try
                {
                    relationships = model.ModelRelationships;
                    int count = relationships.Count;

                    for (int i = 1; i <= count; i++)
                    {
                        dynamic? relationship = null;
                        dynamic? fkColumn = null;
                        dynamic? pkColumn = null;
                        dynamic? fkTable = null;
                        dynamic? pkTable = null;
                        try
                        {
                            relationship = relationships.Item(i);

                            // Get foreign key column and table
                            fkColumn = relationship.ForeignKeyColumn;
                            fkTable = fkColumn.Parent;
                            string fromColumn = fkColumn.Name?.ToString() ?? "";
                            string fromTable = fkTable.Name?.ToString() ?? "";

                            // Get primary key column and table
                            pkColumn = relationship.PrimaryKeyColumn;
                            pkTable = pkColumn.Parent;
                            string toColumn = pkColumn.Name?.ToString() ?? "";
                            string toTable = pkTable.Name?.ToString() ?? "";

                            // Get relationship status
                            bool isActive = relationship.Active ?? true;

                            var relationshipInfo = new DataModelRelationshipInfo
                            {
                                FromTable = fromTable,
                                FromColumn = fromColumn,
                                ToTable = toTable,
                                ToColumn = toColumn,
                                IsActive = isActive
                            };

                            result.Relationships.Add(relationshipInfo);
                        }
                        finally
                        {
                            ComUtilities.Release(ref pkTable);
                            ComUtilities.Release(ref fkTable);
                            ComUtilities.Release(ref pkColumn);
                            ComUtilities.Release(ref fkColumn);
                            ComUtilities.Release(ref relationship);
                        }
                    }

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref relationships);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing relationships: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(IExcelBatch batch, string? tableName = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = tableName != null ? $"model-refresh-table:{tableName}" : "model-refresh"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = ctx.Book.Model;

                if (tableName != null)
                {
                    // Refresh specific table
                    dynamic? table = ComUtilities.FindModelTable(model, tableName);
                    if (table == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Table '{tableName}' not found in Data Model.";
                        return result;
                    }

                    try
                    {
                        table.Refresh();
                        result.Success = true;
                        result.SuggestedNextActions = new List<string>
                        {
                            $"Table '{tableName}' refreshed successfully",
                            "Use 'model-list-tables' to verify record counts"
                        };
                    }
                    finally
                    {
                        ComUtilities.Release(ref table);
                    }
                }
                else
                {
                    // Refresh entire model
                    try
                    {
                        model.Refresh();
                        result.Success = true;
                        result.SuggestedNextActions = new List<string>
                        {
                            "All Data Model tables refreshed successfully",
                            "Use 'model-list-tables' to verify record counts"
                        };
                    }
                    catch (Exception refreshEx)
                    {
                        // Model.Refresh() may not be supported in all Excel versions
                        // Fall back to refreshing tables individually
                        result.ErrorMessage = $"Model-level refresh not supported. Try refreshing tables individually. Error: {refreshEx.Message}";
                        result.Success = false;
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error refreshing Data Model: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteMeasureAsync(IExcelBatch batch, string measureName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-delete-measure"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = ComUtilities.FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    var measureNames = DataModelHelpers.GetModelMeasureNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Measure '{measureName}' not found in Data Model.";

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

                    if (suggestions.Any())
                    {
                        result.SuggestedNextActions = suggestions;
                    }

                    return result;
                }

                // Delete the measure
                measure.Delete();

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Measure '{measureName}' deleted successfully",
                    "Use 'model-list-measures' to verify deletion",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Measure deleted. Next, verify remaining measures or create new ones.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting measure: {ex.Message}";
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
    public async Task<OperationResult> DeleteRelationshipAsync(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-delete-relationship"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? modelRelationships = null;
            dynamic? relationship = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = ctx.Book.Model;
                modelRelationships = model.ModelRelationships;

                // Find the relationship
                bool found = false;
                for (int i = 1; i <= modelRelationships.Count; i++)
                {
                    try
                    {
                        relationship = modelRelationships.Item(i);

                        dynamic? fkColumn = relationship.ForeignKeyColumn;
                        dynamic? pkColumn = relationship.PrimaryKeyColumn;

                        try
                        {
                            dynamic? fkTable = fkColumn.Parent;
                            dynamic? pkTable = pkColumn.Parent;

                            string currentFromTable = fkTable?.Name?.ToString() ?? "";
                            string currentFromColumn = fkColumn?.Name?.ToString() ?? "";
                            string currentToTable = pkTable?.Name?.ToString() ?? "";
                            string currentToColumn = pkColumn?.Name?.ToString() ?? "";

                            ComUtilities.Release(ref fkTable);
                            ComUtilities.Release(ref pkTable);

                            if (currentFromTable.Equals(fromTable, StringComparison.OrdinalIgnoreCase) &&
                                currentFromColumn.Equals(fromColumn, StringComparison.OrdinalIgnoreCase) &&
                                currentToTable.Equals(toTable, StringComparison.OrdinalIgnoreCase) &&
                                currentToColumn.Equals(toColumn, StringComparison.OrdinalIgnoreCase))
                            {
                                // Delete the relationship
                                relationship.Delete();
                                found = true;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref fkColumn);
                            ComUtilities.Release(ref pkColumn);
                        }
                    }
                    finally
                    {
                        if (!found || i < modelRelationships.Count)
                        {
                            ComUtilities.Release(ref relationship);
                        }
                    }
                }

                if (!found)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} not found in Data Model.";
                    result.SuggestedNextActions = new List<string>
                    {
                        "Use 'model-list-relationships' to see available relationships",
                        "Check table and column names for typos",
                        "Verify the relationship exists in the Data Model"
                    };
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully",
                    "Use 'model-list-relationships' to verify deletion",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Relationship deleted. Next, verify remaining relationships or create new ones.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting relationship: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref relationship);
                ComUtilities.Release(ref modelRelationships);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }
}

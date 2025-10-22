using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model management commands - Core data layer (no console output)
/// Provides read-only access to Excel Data Model (PowerPivot) objects
/// </summary>
public class DataModelCommands : IDataModelCommands
{
    /// <inheritdoc />
    public DataModelTableListResult ListTables(string filePath)
    {
        var result = new DataModelTableListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        return WithExcel(filePath, save: false, (excel, workbook) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModel(workbook))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model. Load data to Data Model first using Power Query or external data sources.";
                    return result;
                }

                model = workbook.Model;
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
                            ReleaseComObject(ref table);
                        }
                    }

                    result.Success = true;
                }
                finally
                {
                    ReleaseComObject(ref modelTables);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error accessing Data Model: {ex.Message}";
            }
            finally
            {
                ReleaseComObject(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelMeasureListResult ListMeasures(string filePath, string? tableName = null)
    {
        var result = new DataModelMeasureListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        return WithExcel(filePath, save: false, (excel, workbook) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModel(workbook))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = workbook.Model;
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
                                    ReleaseComObject(ref measure);
                                }
                            }
                        }
                        finally
                        {
                            ReleaseComObject(ref measures);
                            ReleaseComObject(ref table);
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
                    ReleaseComObject(ref modelTables);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing measures: {ex.Message}";
            }
            finally
            {
                ReleaseComObject(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelMeasureViewResult ViewMeasure(string filePath, string measureName)
    {
        var result = new DataModelMeasureViewResult
        {
            FilePath = filePath,
            MeasureName = measureName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        return WithExcel(filePath, save: false, (excel, workbook) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModel(workbook))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = workbook.Model;

                // Find the measure
                measure = FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    var measureNames = GetModelMeasureNames(model);
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
                            ReleaseComObject(ref formatInfo);
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
                ReleaseComObject(ref measure);
                ReleaseComObject(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportMeasure(string filePath, string measureName, string outputFile)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "model-export-measure"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

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

        WithExcel(filePath, save: false, (excel, workbook) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModel(workbook))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return 1;
                }

                model = workbook.Model;

                // Find the measure
                measure = FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Measure '{measureName}' not found in Data Model.";
                    return 1;
                }

                // Get measure details
                string daxFormula = measure.Formula?.ToString() ?? "";
                string? description = measure.Description?.ToString();
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
                            ReleaseComObject(ref formatInfo);
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
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting measure: {ex.Message}";
                return 1;
            }
            finally
            {
                ReleaseComObject(ref measure);
                ReleaseComObject(ref model);
            }
        });

        return await Task.FromResult(result);
    }

    /// <inheritdoc />
    public DataModelRelationshipListResult ListRelationships(string filePath)
    {
        var result = new DataModelRelationshipListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        return WithExcel(filePath, save: false, (excel, workbook) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModel(workbook))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = workbook.Model;
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
                            ReleaseComObject(ref pkTable);
                            ReleaseComObject(ref fkTable);
                            ReleaseComObject(ref pkColumn);
                            ReleaseComObject(ref fkColumn);
                            ReleaseComObject(ref relationship);
                        }
                    }

                    result.Success = true;
                }
                finally
                {
                    ReleaseComObject(ref relationships);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing relationships: {ex.Message}";
            }
            finally
            {
                ReleaseComObject(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public OperationResult Refresh(string filePath, string? tableName = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = tableName != null ? $"model-refresh-table:{tableName}" : "model-refresh"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        return WithExcel(filePath, save: true, (excel, workbook) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModel(workbook))
                {
                    result.Success = false;
                    result.ErrorMessage = "This workbook does not contain a Data Model.";
                    return result;
                }

                model = workbook.Model;

                if (tableName != null)
                {
                    // Refresh specific table
                    dynamic? table = FindModelTable(model, tableName);
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
                        ReleaseComObject(ref table);
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
                ReleaseComObject(ref model);
            }

            return result;
        });
    }
}

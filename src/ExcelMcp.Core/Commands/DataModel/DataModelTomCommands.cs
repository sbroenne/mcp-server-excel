using Microsoft.AnalysisServices.Tabular;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model TOM (Tabular Object Model) commands - Core data layer
/// Provides create and update capabilities using Microsoft Analysis Services TOM API
/// </summary>
public class DataModelTomCommands : IDataModelTomCommands
{
    /// <inheritdoc />
    public OperationResult CreateMeasure(
        string filePath,
        string tableName,
        string measureName,
        string daxFormula,
        string? description = null,
        string? formatString = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-create-measure"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (string.IsNullOrWhiteSpace(measureName))
        {
            result.Success = false;
            result.ErrorMessage = "Measure name cannot be empty";
            return result;
        }

        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            result.Success = false;
            result.ErrorMessage = "DAX formula cannot be empty";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find the table
                var table = TomHelper.FindTable(model, tableName);
                if (table == null)
                {
                    var tableNames = TomHelper.GetTableNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found in Data Model.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'model-list-tables' to see all tables",
                        "Verify table name spelling and case"
                    };
                    return result;
                }

                // Check if measure already exists
                var existingMeasure = table.Measures.FirstOrDefault(m =>
                    m.Name.Equals(measureName, StringComparison.OrdinalIgnoreCase));

                if (existingMeasure != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Measure '{measureName}' already exists in table '{tableName}'.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Use 'tom-update-measure' to modify existing measure",
                        "Choose a different measure name",
                        "Use 'model-list-measures' to see existing measures"
                    };
                    return result;
                }

                // Create new measure
                var newMeasure = new Measure
                {
                    Name = measureName,
                    Expression = daxFormula
                };

                if (!string.IsNullOrWhiteSpace(description))
                {
                    newMeasure.Description = description;
                }

                if (!string.IsNullOrWhiteSpace(formatString))
                {
                    newMeasure.FormatString = formatString;
                }

                table.Measures.Add(newMeasure);

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Measure '{measureName}' created successfully in table '{tableName}'",
                    "Use 'model-view-measure' to verify the DAX formula",
                    "Use 'model-refresh' to update calculations",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Measure created. Next, refresh Data Model to apply calculations.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error creating measure: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify DAX formula syntax is correct",
                "Check that table references exist in the model",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public OperationResult UpdateMeasure(
        string filePath,
        string measureName,
        string? daxFormula = null,
        string? description = null,
        string? formatString = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-update-measure"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (string.IsNullOrWhiteSpace(measureName))
        {
            result.Success = false;
            result.ErrorMessage = "Measure name cannot be empty";
            return result;
        }

        // At least one update parameter must be provided
        if (daxFormula == null && description == null && formatString == null)
        {
            result.Success = false;
            result.ErrorMessage = "At least one property must be specified for update (daxFormula, description, or formatString)";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find the measure
                var measure = TomHelper.FindMeasure(model, measureName);
                if (measure == null)
                {
                    var measureNames = TomHelper.GetMeasureNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Measure '{measureName}' not found in Data Model.";

                    // Suggest similar measure names - filter first, then transform
                    var suggestions = measureNames
                        .Where(m => m.Contains(measureName, StringComparison.OrdinalIgnoreCase))
                        .Take(3)
                        .Select(m => $"Try measure: {m}")
                        .ToList();

                    result.SuggestedNextActions = suggestions.Any()
                        ? suggestions
                        : new List<string> { "Use 'model-list-measures' to see available measures" };

                    return result;
                }

                // Update properties
                bool updated = false;

                if (daxFormula != null && !string.IsNullOrWhiteSpace(daxFormula))
                {
                    measure.Expression = daxFormula;
                    updated = true;
                }

                if (description != null)
                {
                    measure.Description = description;
                    updated = true;
                }

                if (formatString != null)
                {
                    measure.FormatString = formatString;
                    updated = true;
                }

                if (!updated)
                {
                    result.Success = false;
                    result.ErrorMessage = "No valid updates provided";
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Measure '{measureName}' updated successfully",
                    "Use 'model-view-measure' to verify changes",
                    "Use 'model-refresh' to update calculations",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Measure updated. Next, refresh Data Model to apply new calculations.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error updating measure: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify DAX formula syntax is correct",
                "Check that measure exists in the model",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public OperationResult CreateRelationship(
        string filePath,
        string fromTable,
        string fromColumn,
        string toTable,
        string toColumn,
        bool isActive = true,
        string crossFilterDirection = "Single")
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-create-relationship"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        // Validate parameters
        if (string.IsNullOrWhiteSpace(fromTable) || string.IsNullOrWhiteSpace(fromColumn) ||
            string.IsNullOrWhiteSpace(toTable) || string.IsNullOrWhiteSpace(toColumn))
        {
            result.Success = false;
            result.ErrorMessage = "Table and column names cannot be empty";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find tables
                var fromTbl = TomHelper.FindTable(model, fromTable);
                var toTbl = TomHelper.FindTable(model, toTable);

                if (fromTbl == null || toTbl == null)
                {
                    var tableNames = TomHelper.GetTableNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Table not found: {(fromTbl == null ? fromTable : toTable)}";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'model-list-tables' to see all tables"
                    };
                    return result;
                }

                // Find columns
                var fromCol = TomHelper.FindColumn(fromTbl, fromColumn);
                var toCol = TomHelper.FindColumn(toTbl, toColumn);

                if (fromCol == null || toCol == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column not found: {(fromCol == null ? $"{fromTable}.{fromColumn}" : $"{toTable}.{toColumn}")}";
                    result.SuggestedNextActions = new List<string>
                    {
                        "Verify column names are correct",
                        "Use Power Pivot to view available columns"
                    };
                    return result;
                }

                // Check if relationship already exists
                var existing = TomHelper.FindRelationship(model, fromTable, fromColumn, toTable, toColumn);
                if (existing != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} already exists.";
                    result.SuggestedNextActions = new List<string>
                    {
                        "Use 'tom-update-relationship' to modify existing relationship",
                        "Use 'model-list-relationships' to see all relationships"
                    };
                    return result;
                }

                // Parse cross-filter direction
                var crossFilter = crossFilterDirection.Equals("Both", StringComparison.OrdinalIgnoreCase)
                    ? CrossFilteringBehavior.BothDirections
                    : CrossFilteringBehavior.OneDirection;

                // Create relationship
                var relationship = new SingleColumnRelationship
                {
                    Name = $"{fromTable}_{fromColumn}_to_{toTable}_{toColumn}",
                    FromColumn = fromCol,
                    ToColumn = toCol,
                    FromCardinality = RelationshipEndCardinality.Many,
                    ToCardinality = RelationshipEndCardinality.One,
                    IsActive = isActive,
                    CrossFilteringBehavior = crossFilter
                };

                model.Relationships.Add(relationship);

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Relationship created from {fromTable}.{fromColumn} to {toTable}.{toColumn}",
                    "Use 'model-list-relationships' to verify",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Relationship created. Next, verify with list-relationships.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error creating relationship: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify table and column names are correct",
                "Check that columns have compatible data types",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public OperationResult UpdateRelationship(
        string filePath,
        string fromTable,
        string fromColumn,
        string toTable,
        string toColumn,
        bool? isActive = null,
        string? crossFilterDirection = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-update-relationship"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        // At least one update parameter must be provided
        if (isActive == null && crossFilterDirection == null)
        {
            result.Success = false;
            result.ErrorMessage = "At least one property must be specified for update (isActive or crossFilterDirection)";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find relationship
                var relationship = TomHelper.FindRelationship(model, fromTable, fromColumn, toTable, toColumn);

                if (relationship == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} not found.";
                    result.SuggestedNextActions = new List<string>
                    {
                        "Use 'model-list-relationships' to see available relationships",
                        "Verify table and column names are correct"
                    };
                    return result;
                }

                // Update properties
                bool updated = false;

                if (isActive.HasValue)
                {
                    relationship.IsActive = isActive.Value;
                    updated = true;
                }

                if (!string.IsNullOrWhiteSpace(crossFilterDirection))
                {
                    var crossFilter = crossFilterDirection.Equals("Both", StringComparison.OrdinalIgnoreCase)
                        ? CrossFilteringBehavior.BothDirections
                        : CrossFilteringBehavior.OneDirection;

                    relationship.CrossFilteringBehavior = crossFilter;
                    updated = true;
                }

                if (!updated)
                {
                    result.Success = false;
                    result.ErrorMessage = "No valid updates provided";
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} updated",
                    "Use 'model-list-relationships' to verify changes",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Relationship updated. Next, verify with list-relationships.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error updating relationship: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify relationship exists in the model",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public OperationResult CreateCalculatedColumn(
        string filePath,
        string tableName,
        string columnName,
        string daxFormula,
        string? description = null,
        string dataType = "String")
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-create-calculated-column"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            result.Success = false;
            result.ErrorMessage = "Column name cannot be empty";
            return result;
        }

        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            result.Success = false;
            result.ErrorMessage = "DAX formula cannot be empty";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find the table
                var table = TomHelper.FindTable(model, tableName);
                if (table == null)
                {
                    var tableNames = TomHelper.GetTableNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found in Data Model.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'model-list-tables' to see all tables"
                    };
                    return result;
                }

                // Check if column already exists
                var existingColumn = TomHelper.FindColumn(table, columnName);
                if (existingColumn != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' already exists in table '{tableName}'.";
                    result.SuggestedNextActions = new List<string>
                    {
                        "Choose a different column name",
                        "Use Power Pivot to view existing columns"
                    };
                    return result;
                }

                // Parse data type
                var tomDataType = dataType.ToLowerInvariant() switch
                {
                    "integer" or "int" => DataType.Int64,
                    "double" or "decimal" or "number" => DataType.Double,
                    "boolean" or "bool" => DataType.Boolean,
                    "datetime" or "date" => DataType.DateTime,
                    _ => DataType.String
                };

                // Create calculated column
                var newColumn = new CalculatedColumn
                {
                    Name = columnName,
                    Expression = daxFormula,
                    DataType = tomDataType
                };

                if (!string.IsNullOrWhiteSpace(description))
                {
                    newColumn.Description = description;
                }

                table.Columns.Add(newColumn);

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Calculated column '{columnName}' created successfully in table '{tableName}'",
                    "Use 'model-refresh' to populate column values",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Calculated column created. Next, refresh Data Model to populate values.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error creating calculated column: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify DAX formula syntax is correct",
                "Check that column references exist in the table",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public DataModelValidationResult ValidateDax(string filePath, string daxFormula)
    {
        var result = new DataModelValidationResult
        {
            FilePath = filePath,
            DaxFormula = daxFormula
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            result.IsValid = false;
            result.ValidationError = "File not found";
            return result;
        }

        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            result.Success = false;
            result.ErrorMessage = "DAX formula cannot be empty";
            result.IsValid = false;
            result.ValidationError = "DAX formula is empty";
            return result;
        }

        try
        {
            var (isValid, errorMessage) = TomHelper.ValidateDaxFormula(filePath, daxFormula);

            result.Success = true;
            result.IsValid = isValid;
            result.ValidationError = errorMessage;

            // Use ternary operator for conditional assignment
            result.SuggestedNextActions = isValid
                ? new List<string>
                {
                    "DAX formula syntax appears valid",
                    "Create a measure using this formula",
                    "Test the formula with actual data"
                }
                : new List<string>
                {
                    "Review DAX formula syntax",
                    "Check for balanced parentheses and brackets",
                    "Verify table and column references exist"
                };

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error validating DAX: {ex.Message}";
            result.IsValid = false;
            result.ValidationError = ex.Message;
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> ImportMeasures(string filePath, string importFile)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-import-measures"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (!File.Exists(importFile))
        {
            result.Success = false;
            result.ErrorMessage = $"Import file not found: {importFile}";
            return result;
        }

        try
        {
            // Validate import file path
            importFile = PathValidator.ValidateExistingFile(importFile, nameof(importFile));
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid import file path: {ex.Message}";
            return result;
        }

        try
        {
            // Read import file
            var content = await File.ReadAllTextAsync(importFile);

            // For now, support simple DAX file format (measure name and formula)
            // Future enhancement: Support JSON format with multiple measures
            var extension = Path.GetExtension(importFile).ToLowerInvariant();

            if (extension == ".dax")
            {
                // Parse DAX file format
                result.Success = false;
                result.ErrorMessage = "DAX file import not yet implemented. Use JSON format for now.";
                return result;
            }
            else if (extension == ".json")
            {
                result.Success = false;
                result.ErrorMessage = "JSON measure import not yet implemented.";
                return result;
            }
            else
            {
                result.Success = false;
                result.ErrorMessage = $"Unsupported import file format: {extension}. Use .dax or .json";
                return result;
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error importing measures: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public DataModelCalculatedColumnListResult ListCalculatedColumns(string filePath, string? tableName = null)
    {
        var result = new DataModelCalculatedColumnListResult
        {
            FilePath = filePath
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // If table name provided, filter to that table
                var tablesToSearch = string.IsNullOrWhiteSpace(tableName)
                    ? model.Tables.ToList()
                    : new List<Microsoft.AnalysisServices.Tabular.Table>
                    {
                        TomHelper.FindTable(model, tableName) ?? throw new InvalidOperationException($"Table '{tableName}' not found")
                    };

                // Use LINQ Select to transform columns to info objects
                var calculatedColumns = tablesToSearch
                    .SelectMany(table => table.Columns
                        .OfType<Microsoft.AnalysisServices.Tabular.CalculatedColumn>()
                        .Select(column => new DataModelCalculatedColumnInfo
                        {
                            Name = column.Name,
                            Table = table.Name,
                            FormulaPreview = column.Expression?.Length > 60
                                ? column.Expression[..57] + "..."
                                : column.Expression ?? "",
                            DataType = column.DataType.ToString(),
                            Description = column.Description
                        }))
                    .ToList();

                result.CalculatedColumns = calculatedColumns;

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Found {result.CalculatedColumns.Count} calculated column(s)",
                    "Use 'view-column' to see full DAX formula",
                    "Use 'update-column' to modify column properties",
                    "Use 'delete-column' to remove columns"
                };
                result.WorkflowHint = "Calculated columns listed. Next, view or modify column formulas.";

                return result;
            }, saveChanges: false);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error listing calculated columns: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify file has Data Model enabled",
                "Ensure TOM API connection is available",
                "Check file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public DataModelCalculatedColumnViewResult ViewCalculatedColumn(string filePath, string tableName, string columnName)
    {
        var result = new DataModelCalculatedColumnViewResult
        {
            FilePath = filePath,
            TableName = tableName,
            ColumnName = columnName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (string.IsNullOrWhiteSpace(tableName))
        {
            result.Success = false;
            result.ErrorMessage = "Table name cannot be empty";
            return result;
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            result.Success = false;
            result.ErrorMessage = "Column name cannot be empty";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find the table
                var table = TomHelper.FindTable(model, tableName);
                if (table == null)
                {
                    var tableNames = TomHelper.GetTableNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found in Data Model.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'list-tables' to see all tables"
                    };
                    return result;
                }

                // Find the calculated column
                var column = table.Columns.OfType<Microsoft.AnalysisServices.Tabular.CalculatedColumn>()
                    .FirstOrDefault(c => c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Calculated column '{columnName}' not found in table '{tableName}'.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Use 'list-columns' to see columns in table '{tableName}'",
                        "Check column name spelling and case"
                    };
                    return result;
                }

                // Populate result
                result.DaxFormula = column.Expression ?? "";
                result.Description = column.Description;
                result.DataType = column.DataType.ToString();
                result.CharacterCount = result.DaxFormula.Length;
                result.Success = true;

                result.SuggestedNextActions = new List<string>
                {
                    "Use 'update-column' to modify formula or properties",
                    "Use 'delete-column' to remove column",
                    "Use 'model-refresh' to recalculate values"
                };
                result.WorkflowHint = "Column details viewed. Next, update or analyze the DAX formula.";

                return result;
            }, saveChanges: false);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error viewing calculated column: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify table and column names",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public OperationResult UpdateCalculatedColumn(
        string filePath,
        string tableName,
        string columnName,
        string? daxFormula = null,
        string? description = null,
        string? dataType = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-update-calculated-column"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (string.IsNullOrWhiteSpace(tableName))
        {
            result.Success = false;
            result.ErrorMessage = "Table name cannot be empty";
            return result;
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            result.Success = false;
            result.ErrorMessage = "Column name cannot be empty";
            return result;
        }

        // At least one update parameter must be provided
        if (daxFormula == null && description == null && dataType == null)
        {
            result.Success = false;
            result.ErrorMessage = "At least one property must be specified for update (daxFormula, description, or dataType)";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find the table
                var table = TomHelper.FindTable(model, tableName);
                if (table == null)
                {
                    var tableNames = TomHelper.GetTableNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found in Data Model.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'list-tables' to see all tables"
                    };
                    return result;
                }

                // Find the calculated column
                var column = table.Columns.OfType<Microsoft.AnalysisServices.Tabular.CalculatedColumn>()
                    .FirstOrDefault(c => c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Calculated column '{columnName}' not found in table '{tableName}'.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Use 'list-columns' to see columns in table '{tableName}'",
                        "Check column name spelling"
                    };
                    return result;
                }

                // Update properties
                bool updated = false;

                if (daxFormula != null && !string.IsNullOrWhiteSpace(daxFormula))
                {
                    column.Expression = daxFormula;
                    updated = true;
                }

                if (description != null)
                {
                    column.Description = description;
                    updated = true;
                }

                if (dataType != null && !string.IsNullOrWhiteSpace(dataType))
                {
                    var tomDataType = dataType.ToLowerInvariant() switch
                    {
                        "integer" or "int" => Microsoft.AnalysisServices.Tabular.DataType.Int64,
                        "double" or "decimal" or "number" => Microsoft.AnalysisServices.Tabular.DataType.Double,
                        "boolean" or "bool" => Microsoft.AnalysisServices.Tabular.DataType.Boolean,
                        "datetime" or "date" => Microsoft.AnalysisServices.Tabular.DataType.DateTime,
                        _ => Microsoft.AnalysisServices.Tabular.DataType.String
                    };
                    column.DataType = tomDataType;
                    updated = true;
                }

                if (!updated)
                {
                    result.Success = false;
                    result.ErrorMessage = "No valid updates provided";
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Calculated column '{columnName}' in table '{tableName}' updated successfully",
                    "Use 'view-column' to verify changes",
                    "Use 'model-refresh' to recalculate values",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Column updated. Next, refresh Data Model to recalculate values.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error updating calculated column: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify DAX formula syntax is correct",
                "Check that column exists in the table",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }

    /// <inheritdoc />
    public OperationResult DeleteCalculatedColumn(string filePath, string tableName, string columnName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "tom-delete-calculated-column"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (string.IsNullOrWhiteSpace(tableName))
        {
            result.Success = false;
            result.ErrorMessage = "Table name cannot be empty";
            return result;
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            result.Success = false;
            result.ErrorMessage = "Column name cannot be empty";
            return result;
        }

        try
        {
            return TomHelper.WithTomServer(filePath, (server, model) =>
            {
                // Find the table
                var table = TomHelper.FindTable(model, tableName);
                if (table == null)
                {
                    var tableNames = TomHelper.GetTableNames(model);
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found in Data Model.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'list-tables' to see all tables"
                    };
                    return result;
                }

                // Find the calculated column
                var column = table.Columns.OfType<Microsoft.AnalysisServices.Tabular.CalculatedColumn>()
                    .FirstOrDefault(c => c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Calculated column '{columnName}' not found in table '{tableName}'.";
                    result.SuggestedNextActions = new List<string>
                    {
                        $"Use 'list-columns' to see columns in table '{tableName}'",
                        "Check column name spelling"
                    };
                    return result;
                }

                // Delete the column
                table.Columns.Remove(column);

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Calculated column '{columnName}' deleted from table '{tableName}'",
                    "Use 'list-columns' to verify deletion",
                    "Changes saved to workbook"
                };
                result.WorkflowHint = "Column deleted. Next, verify remaining columns or refresh Data Model.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error deleting calculated column: {ex.Message}";
            result.SuggestedNextActions = new List<string>
            {
                "Verify column exists in the table",
                "Check that column is not referenced by measures or other columns",
                "Ensure file is not locked by Excel"
            };
            return result;
        }
    }
}

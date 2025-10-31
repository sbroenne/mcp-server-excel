using Microsoft.AnalysisServices.Tabular;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model TOM (Tabular Object Model) commands - Calculated column operations
/// Provides calculated column support using Microsoft Analysis Services TOM API
/// Note: Measure and relationship operations use COM API (see DataModelCommands)
/// </summary>
public class DataModelTomCommands : IDataModelTomCommands
{
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
                    result.SuggestedNextActions =
                    [
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'model-list-tables' to see all tables"
                    ];
                    return result;
                }

                // Check if column already exists
                var existingColumn = TomHelper.FindColumn(table, columnName);
                if (existingColumn != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' already exists in table '{tableName}'.";
                    result.SuggestedNextActions =
                    [
                        "Choose a different column name",
                        "Use Power Pivot to view existing columns"
                    ];
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
                result.SuggestedNextActions =
                [
                    $"Calculated column '{columnName}' created successfully in table '{tableName}'",
                    "Use 'model-refresh' to populate column values",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Calculated column created. Next, refresh Data Model to populate values.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error creating calculated column: {ex.Message}";
            result.SuggestedNextActions =
            [
                "Verify DAX formula syntax is correct",
                "Check that column references exist in the table",
                "Ensure file is not locked by Excel"
            ];
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
                ?
                [
                    "DAX formula syntax appears valid",
                    "Create a measure using this formula",
                    "Test the formula with actual data"
                ]
                :
                [
                    "Review DAX formula syntax",
                    "Check for balanced parentheses and brackets",
                    "Verify table and column references exist"
                ];

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
                    :
                    [
                        TomHelper.FindTable(model, tableName) ?? throw new InvalidOperationException($"Table '{tableName}' not found")
                    ];

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
                result.SuggestedNextActions =
                [
                    $"Found {result.CalculatedColumns.Count} calculated column(s)",
                    "Use 'view-column' to see full DAX formula",
                    "Use 'update-column' to modify column properties",
                    "Use 'delete-column' to remove columns"
                ];
                result.WorkflowHint = "Calculated columns listed. Next, view or modify column formulas.";

                return result;
            }, saveChanges: false);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error listing calculated columns: {ex.Message}";
            result.SuggestedNextActions =
            [
                "Verify file has Data Model enabled",
                "Ensure TOM API connection is available",
                "Check file is not locked by Excel"
            ];
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
                    result.SuggestedNextActions =
                    [
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'list-tables' to see all tables"
                    ];
                    return result;
                }

                // Find the calculated column
                var column = table.Columns.OfType<Microsoft.AnalysisServices.Tabular.CalculatedColumn>()
                    .FirstOrDefault(c => c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Calculated column '{columnName}' not found in table '{tableName}'.";
                    result.SuggestedNextActions =
                    [
                        $"Use 'list-columns' to see columns in table '{tableName}'",
                        "Check column name spelling and case"
                    ];
                    return result;
                }

                // Populate result
                result.DaxFormula = column.Expression ?? "";
                result.Description = column.Description;
                result.DataType = column.DataType.ToString();
                result.CharacterCount = result.DaxFormula.Length;
                result.Success = true;

                result.SuggestedNextActions =
                [
                    "Use 'update-column' to modify formula or properties",
                    "Use 'delete-column' to remove column",
                    "Use 'model-refresh' to recalculate values"
                ];
                result.WorkflowHint = "Column details viewed. Next, update or analyze the DAX formula.";

                return result;
            }, saveChanges: false);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error viewing calculated column: {ex.Message}";
            result.SuggestedNextActions =
            [
                "Verify table and column names",
                "Ensure file is not locked by Excel"
            ];
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
                    result.SuggestedNextActions =
                    [
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'list-tables' to see all tables"
                    ];
                    return result;
                }

                // Find the calculated column
                var column = table.Columns.OfType<Microsoft.AnalysisServices.Tabular.CalculatedColumn>()
                    .FirstOrDefault(c => c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Calculated column '{columnName}' not found in table '{tableName}'.";
                    result.SuggestedNextActions =
                    [
                        $"Use 'list-columns' to see columns in table '{tableName}'",
                        "Check column name spelling"
                    ];
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
                result.SuggestedNextActions =
                [
                    $"Calculated column '{columnName}' in table '{tableName}' updated successfully",
                    "Use 'view-column' to verify changes",
                    "Use 'model-refresh' to recalculate values",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Column updated. Next, refresh Data Model to recalculate values.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error updating calculated column: {ex.Message}";
            result.SuggestedNextActions =
            [
                "Verify DAX formula syntax is correct",
                "Check that column exists in the table",
                "Ensure file is not locked by Excel"
            ];
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
                    result.SuggestedNextActions =
                    [
                        $"Available tables: {string.Join(", ", tableNames)}",
                        "Use 'list-tables' to see all tables"
                    ];
                    return result;
                }

                // Find the calculated column
                var column = table.Columns.OfType<Microsoft.AnalysisServices.Tabular.CalculatedColumn>()
                    .FirstOrDefault(c => c.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Calculated column '{columnName}' not found in table '{tableName}'.";
                    result.SuggestedNextActions =
                    [
                        $"Use 'list-columns' to see columns in table '{tableName}'",
                        "Check column name spelling"
                    ];
                    return result;
                }

                // Delete the column
                table.Columns.Remove(column);

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Calculated column '{columnName}' deleted from table '{tableName}'",
                    "Use 'list-columns' to verify deletion",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Column deleted. Next, verify remaining columns or refresh Data Model.";

                return result;
            }, saveChanges: true);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error deleting calculated column: {ex.Message}";
            result.SuggestedNextActions =
            [
                "Verify column exists in the table",
                "Check that column is not referenced by measures or other columns",
                "Ensure file is not locked by Excel"
            ];
            return result;
        }
    }
}

namespace Sbroenne.ExcelMcp.Core.Commands.DataModel;

/// <summary>
/// Provides workflow guidance for LLMs working with Data Model operations.
/// Suggests next actions based on operation context to improve LLM effectiveness.
/// Includes batch mode suggestions for multi-operation workflows.
/// </summary>
public static class DataModelWorkflowGuidance
{
    /// <summary>
    /// Get suggested next steps after creating a DAX measure
    /// </summary>
    /// <param name="success">Whether the measure creation succeeded</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterCreateMeasure(bool success = true, bool usedBatchMode = false)
    {
        if (!success)
        {
            return new List<string>
            {
                "Measure creation failed",
                "Review error message for DAX formula issues",
                "Verify table name and measure name are valid",
                "Check DAX syntax and referenced columns exist"
            };
        }

        var suggestions = new List<string>
        {
            "Measure created successfully in Data Model",
            "Use 'dm-list-measures' to see all measures",
            "Use 'dm-view-measure' to inspect DAX formula",
            "Measure is now available in PivotTables and Power BI"
        };

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "Creating multiple measures? Use begin_excel_batch to keep Data Model open (much faster)");
        }

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps after creating a relationship
    /// </summary>
    /// <param name="success">Whether the relationship creation succeeded</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterCreateRelationship(bool success = true, bool usedBatchMode = false)
    {
        if (!success)
        {
            return new List<string>
            {
                "Relationship creation failed",
                "Verify both tables and columns exist in Data Model",
                "Check column data types are compatible",
                "Ensure relationship doesn't create circular dependencies"
            };
        }

        var suggestions = new List<string>
        {
            "Relationship created successfully",
            "Use 'dm-list-relationships' to see all relationships",
            "Use 'dm-refresh' to validate relationship with data",
            "Relationship enables cross-table DAX calculations"
        };

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "Creating multiple relationships? Use begin_excel_batch to keep Data Model open");
        }

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps after creating a calculated column
    /// </summary>
    /// <param name="success">Whether the column creation succeeded</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterCreateColumn(bool success = true, bool usedBatchMode = false)
    {
        if (!success)
        {
            return new List<string>
            {
                "Calculated column creation failed",
                "Review DAX formula for syntax errors",
                "Verify referenced columns exist",
                "Check data types are compatible for calculations"
            };
        }

        var suggestions = new List<string>
        {
            "Calculated column created successfully",
            "Column is available in table for measures and relationships",
            "Use 'refresh-datamodel' to populate column with calculated values",
            "Consider using measures instead of columns for better performance"
        };

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "Creating multiple columns? Use begin_excel_batch to keep Data Model open");
        }

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps after listing Data Model objects
    /// </summary>
    /// <param name="objectType">Type of objects listed (tables, measures, relationships)</param>
    /// <param name="count">Number of objects found</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterList(string objectType, int count)
    {
        var suggestions = new List<string>();

        if (count == 0)
        {
            suggestions.Add($"No {objectType} found in Data Model");
            suggestions.Add("Use Power Query to load tables to Data Model");
            suggestions.Add("Use 'powerquery set-load-to-data-model' to configure queries");
        }
        else
        {
            suggestions.Add($"Found {count} {objectType} in Data Model");

            if (objectType.Contains("measure", StringComparison.OrdinalIgnoreCase))
            {
                suggestions.Add("Use 'dm-view-measure' to inspect DAX formulas");
                suggestions.Add("Use 'dm-update-measure' to modify existing measures");
            }
            else if (objectType.Contains("table", StringComparison.OrdinalIgnoreCase))
            {
                suggestions.Add("Use 'dm-list-relationships' to see how tables are connected");
                suggestions.Add("Use 'dm-create-measure' to add calculations");
            }
            else if (objectType.Contains("relationship", StringComparison.OrdinalIgnoreCase))
            {
                suggestions.Add("Relationships enable cross-table DAX calculations");
                suggestions.Add("Use 'dm-create-measure' to leverage relationships");
            }
        }

        return suggestions;
    }

    /// <summary>
    /// Get workflow hint based on operation context
    /// </summary>
    /// <param name="operation">The operation being performed</param>
    /// <param name="success">Whether the operation succeeded</param>
    /// <param name="usedBatchMode">Whether batch mode was used</param>
    /// <returns>Contextual workflow hint</returns>
    public static string GetWorkflowHint(string operation, bool success, bool usedBatchMode = false)
    {
        if (!success)
        {
            return $"{operation} failed. Review error details and suggested recovery steps.";
        }

        var hint = operation switch
        {
            "create-measure" => "WORKFLOW: Create Measure → Verify → Use in PivotTable",
            "update-measure" => "WORKFLOW: Update Measure → Refresh → Verify Changes",
            "delete-measure" => "WORKFLOW: Measure deleted, PivotTables may need updates",
            "create-relationship" => "WORKFLOW: Create Relationship → Refresh → Verify Data Flow",
            "create-column" => "WORKFLOW: Create Column → Refresh → Column Available",
            "list-measures" => "WORKFLOW: List → View Details → Create/Update as needed",
            "list-tables" => "WORKFLOW: List Tables → Create Relationships → Add Measures",
            "refresh-datamodel" => "WORKFLOW: Refresh validates all DAX formulas and loads data",
            _ => $"{operation} completed successfully"
        };

        if (!usedBatchMode && IsMultiOperationScenario(operation))
        {
            hint += ". Consider using batch mode for multiple operations.";
        }

        return hint;
    }

    /// <summary>
    /// Check if operation is typically part of multi-operation workflow
    /// </summary>
    private static bool IsMultiOperationScenario(string operation)
    {
        return operation switch
        {
            "create-measure" => true,
            "update-measure" => true,
            "create-relationship" => true,
            "create-column" => true,
            _ => false
        };
    }

    /// <summary>
    /// Get error recovery steps based on error type
    /// </summary>
    /// <param name="errorType">Type of error encountered</param>
    /// <returns>List of recovery steps</returns>
    public static List<string> GetErrorRecoverySteps(string errorType)
    {
        return errorType switch
        {
            "DAXSyntax" => new List<string>
            {
                "Check DAX formula syntax",
                "Verify all function names are spelled correctly",
                "Ensure proper use of commas, parentheses, and brackets",
                "Test formula in smaller parts to isolate error"
            },
            "ColumnNotFound" => new List<string>
            {
                "Verify column name and table name are correct",
                "Check if column exists in the table",
                "Use 'list-tables' to see available columns",
                "Ensure column name follows 'TableName[ColumnName]' format"
            },
            "CircularDependency" => new List<string>
            {
                "Review relationship directions",
                "Check if relationship creates a loop in data model",
                "Consider using different relationship approach",
                "Use inactive relationships with USERELATIONSHIP function"
            },
            "DataTypeIncompatible" => new List<string>
            {
                "Check that columns have compatible data types",
                "Verify numeric columns for relationships are same type",
                "Use FORMAT or VALUE functions to convert types",
                "Review source data in Power Query"
            },
            _ => new List<string>
            {
                "Review error message for specific details",
                "Check Data Model structure and relationships",
                "Verify DAX formulas and column references",
                "Consider simplifying operation to isolate issue"
            }
        };
    }
}

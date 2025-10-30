namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Provides workflow guidance for LLMs working with Worksheet operations.
/// Suggests next actions based on operation context to improve LLM effectiveness.
/// Includes batch mode suggestions for multi-operation workflows.
/// </summary>
public static class WorksheetWorkflowGuidance
{
    /// <summary>
    /// Get suggested next steps after creating a worksheet
    /// </summary>
    /// <param name="success">Whether the worksheet creation succeeded</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterCreate(bool success = true, bool usedBatchMode = false)
    {
        if (!success)
        {
            return
            [
                "Worksheet creation failed",
                "Check that worksheet name is valid (no special characters: \\ / ? * [ ])",
                "Verify worksheet name doesn't already exist",
                "Ensure workbook is not protected"
            ];
        }

        var suggestions = new List<string>
        {
            "Worksheet created successfully",
            "Use 'range-set-values' to add data to the new sheet",
            "Use 'create-table' to structure data as Excel Table",
            "Use 'set-named-range' to create parameter references"
        };

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "Creating multiple sheets? Use begin_excel_batch for complete workbook setup");
        }

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps after renaming a worksheet
    /// </summary>
    /// <param name="success">Whether the rename succeeded</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterRename(bool success = true)
    {
        if (!success)
        {
            return
            [
                "Worksheet rename failed",
                "Check that new name doesn't already exist",
                "Verify name is valid (no special characters: \\ / ? * [ ])",
                "Ensure worksheet is not protected"
            ];
        }

        return
        [
            "Worksheet renamed successfully",
            "Update any formulas or references using the old sheet name",
            "Update Power Query expressions if they reference this sheet",
            "Named ranges referencing this sheet are automatically updated"
        ];
    }

    /// <summary>
    /// Get suggested next steps after copying a worksheet
    /// </summary>
    /// <param name="success">Whether the copy succeeded</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterCopy(bool success = true, bool usedBatchMode = false)
    {
        if (!success)
        {
            return
            [
                "Worksheet copy failed",
                "Check that source worksheet exists",
                "Verify new name doesn't already exist",
                "Ensure sufficient memory for copy operation"
            ];
        }

        var suggestions = new List<string>
        {
            "Worksheet copied successfully",
            "Copy includes all data, formulas, and formatting",
            "Named ranges are copied but may need adjustment",
            "Review formulas to ensure they reference correct sheets"
        };

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "Copying multiple sheets? Use begin_excel_batch to group operations");
        }

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps for setup workflows
    /// </summary>
    /// <param name="operationCount">Number of operations in the setup</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsForSetupWorkflow(int operationCount = 0)
    {
        var suggestions = new List<string>();

        if (operationCount > 1)
        {
            suggestions.Add("Use begin_excel_batch for complete setup (sheets + named ranges + queries)");
            suggestions.Add("Batch mode keeps workbook open across all operations (much faster)");
        }

        suggestions.Add("Create sheets first, then add data and formatting");
        suggestions.Add("Use create-table to structure data");
        suggestions.Add("Add named ranges for parameters");
        suggestions.Add("Import Power Queries to load external data");

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
            "create" => "WORKFLOW: Create Sheet → Add Data → Format → Add Tables/Ranges",
            "rename" => "WORKFLOW: Renamed. Update references in formulas and queries.",
            "copy" => "WORKFLOW: Sheet copied. Review formulas and adjust as needed.",
            "delete" => "WORKFLOW: Sheet deleted. Update dependent formulas and queries.",
            "list" => "WORKFLOW: Review sheets → Plan structure → Create/modify as needed",
            _ => $"{operation} completed successfully"
        };

        if (!usedBatchMode && IsMultiOperationScenario(operation))
        {
            hint += ". Consider using batch mode for multiple sheet operations.";
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
            "create" => true,
            "copy" => true,
            _ => false
        };
    }
}

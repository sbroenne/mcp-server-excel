namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Provides workflow guidance for LLMs working with Power Query operations.
/// Suggests next actions based on operation context to improve LLM effectiveness.
/// </summary>
public static class PowerQueryWorkflowGuidance
{
    /// <summary>
    /// Get suggested next steps after importing a query
    /// </summary>
    /// <param name="isConnectionOnly">Whether the query is connection-only</param>
    /// <param name="hasErrors">Whether errors were detected during import</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterImport(bool isConnectionOnly, bool hasErrors = false, bool usedBatchMode = false)
    {
        if (hasErrors)
        {
            return
            [
                "Query import failed validation",
                "Review error messages and fix issues",
                "Use 'view' to examine M code",
                "Check data source connections and credentials"
            ];
        }

        var suggestions = new List<string>();

        if (isConnectionOnly)
        {
            suggestions.Add("Query imported as connection-only (no data loaded yet)");
            suggestions.Add("Use 'set-load-to-table' with targetSheet parameter to load data to worksheet");
            suggestions.Add("Or use 'set-load-to-data-model' to load to PowerPivot");
            suggestions.Add("Loading will validate the query and enable refresh");

            // Suggest batch mode only if planning multiple operations (load + refresh)
            if (!usedBatchMode)
            {
                suggestions.Add("Planning to configure load AND refresh? Use begin_excel_batch to combine operations");
            }
        }
        else
        {
            suggestions.Add("Query imported and data loaded successfully");
            suggestions.Add("Use 'view' to review M code if needed");
            suggestions.Add("Use 'get-load-config' to check configuration");
        }

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps after updating a query
    /// </summary>
    /// <param name="configPreserved">Whether load configuration was preserved</param>
    /// <param name="hasErrors">Whether errors were detected during update</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterUpdate(bool configPreserved = true, bool hasErrors = false, bool usedBatchMode = false)
    {
        if (hasErrors)
        {
            return
            [
                "Query update failed validation",
                "Review error messages and fix M code issues",
                "Use 'view' to examine updated M code",
                "Revert changes if needed with 'update' using previous version"
            ];
        }

        var suggestions = new List<string>();

        if (configPreserved)
        {
            suggestions.Add("Query updated successfully, load configuration preserved");
            suggestions.Add("Data automatically refreshed with new M code");
            suggestions.Add("Use 'get-load-config' to verify configuration if needed");
        }
        else
        {
            suggestions.Add("Query updated successfully");
            suggestions.Add("Check 'get-load-config' to see if query is loaded anywhere");
            suggestions.Add("Use 'set-load-to-table' to load data if connection-only");
        }

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps after configuring load destination
    /// </summary>
    /// <param name="loadMode">The load mode that was configured</param>
    /// <param name="usedBatchMode">Whether batch mode was used for this operation</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterLoadConfig(string loadMode, bool usedBatchMode = false)
    {
        var suggestions = new List<string>();

        // IMPORTANT: Clarify what load mode means for Data Model workflows
        if (loadMode == "LoadToTable")
        {
            suggestions.Add("Query data loaded to worksheet (visible to users as formatted table)");
            suggestions.Add("IMPORTANT: This is NOT loaded to Power Pivot Data Model yet");
            suggestions.Add("To add to Data Model: Use 'set-load-to-data-model' action (simplest)");
            suggestions.Add("Alternative: Create Excel Table from range using 'excel_table create', then 'add-to-datamodel'");
        }
        else if (loadMode == "LoadToDataModel")
        {
            suggestions.Add("Query data loaded to Power Pivot Data Model (ready for DAX)");
            suggestions.Add("Use 'excel_datamodel' tool for DAX measures and relationships");
            suggestions.Add("Data is in model but NOT visible in worksheet (connection-only to Data Model)");
        }
        else if (loadMode == "LoadToBoth")
        {
            suggestions.Add("Query data loaded to BOTH worksheet AND Power Pivot Data Model");
            suggestions.Add("Data visible in worksheet AND available for DAX measures/relationships");
            suggestions.Add("Use 'excel_datamodel' tool for DAX operations");
        }
        else
        {
            suggestions.Add($"Query configured to load data as: {loadMode}");
        }

        suggestions.Add("Use 'refresh' to reload when source data changes");
        suggestions.Add("Use 'view' to review M code if needed");

        return suggestions;
    }

    /// <summary>
    /// Get suggested next steps after refreshing a query
    /// </summary>
    /// <param name="hasErrors">Whether errors were detected during refresh</param>
    /// <param name="isConnectionOnly">Whether this is a connection-only query</param>
    /// <returns>List of suggested actions for LLM</returns>
    public static List<string> GetNextStepsAfterRefresh(bool hasErrors, bool isConnectionOnly)
    {
        if (hasErrors)
        {
            return
            [
                "Refresh failed - query has errors",
                "Review error messages for specific issues",
                "Use 'view' to examine M code",
                "Common issues: authentication, connectivity, M syntax, privacy levels"
            ];
        }

        if (isConnectionOnly)
        {
            return
            [
                "Query validated successfully (connection-only mode)",
                "No data loaded - query is set to connection-only",
                "Use 'set-load-to-table' to load data to a worksheet",
                "Or use 'set-load-to-data-model' for PowerPivot"
            ];
        }

        return
        [
            "Data refreshed successfully",
            "Query is working correctly",
            "Use 'get-load-config' to check where data is loaded",
            "Data model is up to date"
        ];
    }

    /// <summary>
    /// Get suggested next steps for error recovery scenarios
    /// </summary>
    /// <param name="errorCategory">Category of error (Authentication, Connectivity, Privacy, Syntax, Permissions, Unknown)</param>
    /// <returns>List of recovery steps specific to error type</returns>
    public static List<string> GetErrorRecoverySteps(string errorCategory)
    {
        return errorCategory switch
        {
            "Authentication" =>
            [
                "LLM-Actionable Steps:",
                "  • Verify credentials are configured in Excel for this data source",
                "  • Review M code for authentication method used",
                "Requires User Intervention:",
                "  • Configure data source credentials in Excel (Data → Queries & Connections → Edit)",
                "  • Verify username and password are correct",
                "  • Check if data source requires Windows Authentication vs. Database credentials",
                "  • Confirm service account has necessary permissions"
            ],
            "Connectivity" =>
            [
                "LLM-Actionable Steps:",
                "  • Verify data source URL or path in M code is correct",
                "  • Retry connection with 'pq-refresh' command",
                "Requires User Intervention:",
                "  • Verify network connectivity to data source",
                "  • Check Windows firewall and proxy settings",
                "  • Test connection manually in Excel UI (Data → Queries & Connections → Edit)"
            ],
            "Privacy" =>
            [
                "Formula.Firewall error detected (Power Query privacy levels)",
                "⚠️ Cannot be automated - requires Excel UI configuration",
                "User Action Required:",
                "  1. Open Excel → Data → Queries & Connections",
                "  2. Right-click query → Properties → Privacy Level",
                "  3. Set to Private (recommended) or Organizational/Public",
                "Alternative: Rewrite M code to avoid combining protected sources"
            ],
            "Syntax" =>
            [
                "M code syntax error detected",
                "Use 'view' to examine query formula",
                "Check for missing commas, parentheses, or quotes",
                "Validate function names and parameters"
            ],
            "Permissions" =>
            [
                "LLM-Actionable Steps:",
                "  • Verify file path in M code is accessible",
                "  • Check if data source requires authentication",
                "Requires User Intervention:",
                "  • Check file or data source permissions",
                "  • Verify user has read access to data source",
                "  • Review folder permissions if using file sources",
                "  • Contact administrator if needed"
            ],
            _ =>
            [
                "Review error message for details",
                "Use 'view' to examine M code",
                "Check Excel query settings",
                "Consider testing with simplified query first"
            ]
        };
    }

    /// <summary>
    /// Get workflow hint based on operation context
    /// </summary>
    /// <param name="operation">The operation being performed</param>
    /// <param name="success">Whether the operation succeeded</param>
    /// <returns>Contextual workflow hint</returns>
    public static string GetWorkflowHint(string operation, bool success)
    {
        if (!success)
        {
            return $"{operation} failed. Review error details and suggested recovery steps.";
        }

        return operation switch
        {
            "pq-import" => "Query imported (auto-loads to worksheet unless connection-only specified)",
            "pq-update" => "M code updated, configuration preserved, data NOT refreshed automatically",
            "pq-refresh" => "Query refreshed - loaded latest data from source",
            "pq-set-load-to-table" => "Data loaded to worksheet - use 'pq-refresh' to reload when source changes",
            "pq-set-load-to-data-model" => "Data loaded to PowerPivot - use 'pq-refresh' to reload when source changes",
            "pq-set-load-to-both" => "Data loaded to both destinations - use 'pq-refresh' to reload when source changes",
            "pq-set-connection-only" => "Query set to connection-only (no data loading on refresh)",
            "pq-delete" => "Query removed from workbook (data in worksheets may persist)",
            "pq-export" => "M code exported for version control or documentation",
            "pq-view" => "Review M code before making changes",
            "pq-list" => "Displays all Power Query queries in workbook",
            _ => $"{operation} completed successfully"
        };
    }
}

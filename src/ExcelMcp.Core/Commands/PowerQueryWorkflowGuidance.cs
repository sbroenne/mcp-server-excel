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
            suggestions.Add("Then use 'refresh' to validate the query works");
        }
        else
        {
            suggestions.Add("Query imported and data loaded successfully");
            suggestions.Add("Use 'view' to review M code if needed");
            suggestions.Add("Use 'get-load-config' to check configuration");
        }

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "For multiple imports: Use begin_excel_batch to group operations efficiently");
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
            suggestions.Add("Use 'refresh' to reload data with updated M code");
            suggestions.Add("Check 'get-load-config' to verify load settings");
        }

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "For multiple updates: Use begin_excel_batch to group operations efficiently");
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
        var suggestions = new List<string>
        {
            $"Query configured to load data as: {loadMode}",
            "Use 'refresh' to load data to configured destination",
            "Use 'view' to review M code if needed",
            "Data will now refresh to this location automatically"
        };

        // Add batch mode suggestion if not already using it
        if (!usedBatchMode)
        {
            suggestions.Insert(0, "For configuring multiple queries: Use begin_excel_batch to group operations efficiently");
        }

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
                "Check data source credentials",
                "Verify authentication method in M code",
                "Update connection strings with valid credentials",
                "Consider using Excel credential manager"
            ],
            "Connectivity" =>
            [
                "Verify network connectivity to data source",
                "Check firewall and proxy settings",
                "Confirm data source URL or path is correct",
                "Test connection from Excel manually"
            ],
            "Privacy" =>
            [
                "Privacy level mismatch detected",
                "Use 'update' with privacyLevel parameter (Private, Organizational, or Public)",
                "Review M code for data source combinations",
                "Consider using consistent privacy levels across queries"
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
                "Check file or data source permissions",
                "Verify user has read access to data source",
                "Review folder permissions if using file sources",
                "Contact administrator if needed"
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
            "pq-import" => "WORKFLOW: Import → Configure Load → Refresh → Validate",
            "pq-update" => "WORKFLOW: Update → Auto-Refresh → Verify (config preserved)",
            "pq-refresh" => "WORKFLOW: Refresh validates query and loads latest data",
            "pq-set-load-to-table" => "WORKFLOW: Configure → Refresh → Data appears in worksheet",
            "pq-set-load-to-data-model" => "WORKFLOW: Configure → Refresh → Data available in PowerPivot",
            "pq-set-load-to-both" => "WORKFLOW: Configure → Refresh → Data in worksheet and PowerPivot",
            "pq-set-connection-only" => "WORKFLOW: Query set to connection-only (no data loading)",
            "pq-delete" => "WORKFLOW: Query removed, but data in worksheets may persist",
            "pq-export" => "WORKFLOW: M code exported for version control",
            "pq-view" => "WORKFLOW: Review M code before making changes",
            "pq-list" => "WORKFLOW: List queries to understand workbook structure",
            _ => $"{operation} completed successfully"
        };
    }
}

using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Security;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query management commands - Core data layer (no console output)
/// </summary>
public class PowerQueryCommands : IPowerQueryCommands
{
    private readonly IDataModelCommands _dataModelCommands;

    /// <summary>
    /// Constructor with dependency injection for atomic Data Model operations
    /// </summary>
    /// <param name="dataModelCommands">Data Model commands for atomic refresh operations in SetLoadToDataModelAsync</param>
    public PowerQueryCommands(IDataModelCommands dataModelCommands)
    {
        _dataModelCommands = dataModelCommands ?? throw new ArgumentNullException(nameof(dataModelCommands));
    }

    /// <summary>
    /// Finds the closest matching string using simple Levenshtein distance
    /// </summary>
    private static string? FindClosestMatch(string target, List<string> candidates)
    {
        if (candidates.Count == 0) return null;

        int minDistance = int.MaxValue;
        string? bestMatch = null;

        foreach (var candidate in candidates)
        {
            int distance = ComputeLevenshteinDistance(target.ToLowerInvariant(), candidate.ToLowerInvariant());
            if (distance < minDistance && distance <= Math.Max(target.Length, candidate.Length) / 2)
            {
                minDistance = distance;
                bestMatch = candidate;
            }
        }

        return bestMatch;
    }

    /// <summary>
    /// Computes Levenshtein distance between two strings
    /// </summary>
    private static int ComputeLevenshteinDistance(string s1, string s2)
    {
        int[,] d = new int[s1.Length + 1, s2.Length + 1];

        for (int i = 0; i <= s1.Length; i++)
            d[i, 0] = i;
        for (int j = 0; j <= s2.Length; j++)
            d[0, j] = j;

        for (int i = 1; i <= s1.Length; i++)
        {
            for (int j = 1; j <= s2.Length; j++)
            {
                int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
            }
        }

        return d[s1.Length, s2.Length];
    }

    /// <summary>
    /// Parse COM exception to extract user-friendly Power Query error message
    /// </summary>
    private static string ParsePowerQueryError(COMException comEx)
    {
        var message = comEx.Message;

        if (message.Contains("authentication", StringComparison.OrdinalIgnoreCase))
            return "Data source authentication failed. Check credentials and permissions.";
        if (message.Contains("could not reach", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("unable to connect", StringComparison.OrdinalIgnoreCase))
            return "Cannot connect to data source. Check network connectivity.";
        if (message.Contains("privacy level", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("combine data", StringComparison.OrdinalIgnoreCase))
            return "Formula.Firewall error - privacy levels must be configured in Excel UI (cannot be automated)";
        if (message.Contains("syntax", StringComparison.OrdinalIgnoreCase))
            return "M code syntax error. Review query formula.";
        if (message.Contains("permission", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("access denied", StringComparison.OrdinalIgnoreCase))
            return "Access denied. Check file or data source permissions.";

        return message; // Return original if no pattern matches
    }

    /// <summary>
    /// Extracts file path from File.Contents() in M code
    /// </summary>
    private static string? ExtractFileContentsPath(string mCode)
    {
        // Parse: File.Contents("D:\path\to\file.xlsx")
        // Also handles: File.Contents( "path" ) with optional whitespace
        var match = System.Text.RegularExpressions.Regex.Match(
            mCode, 
            @"File\.Contents\s*\(\s*""([^""]+)""\s*\)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        return match.Success ? match.Groups[1].Value : null;
    }

    /// <summary>
    /// Categorize error type from COM exception
    /// </summary>
    private static string CategorizeError(COMException comEx)
    {
        var message = comEx.Message.ToLower();
        if (message.Contains("authentication")) return "Authentication";
        if (message.Contains("connection") || message.Contains("reach") || message.Contains("connect")) return "Connectivity";
        if (message.Contains("privacy") || message.Contains("combine data")) return "Privacy";
        if (message.Contains("syntax")) return "Syntax";
        if (message.Contains("permission") || message.Contains("access")) return "Permissions";
        return "Unknown";
    }

    /// <summary>
    /// Determine which worksheet a query is loaded to (if any)
    /// </summary>
    private static string? DetermineLoadedSheet(dynamic workbook, string queryName)
    {
        dynamic? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic? worksheet = null;
                dynamic? queryTables = null;
                try
                {
                    worksheet = worksheets.Item(ws);
                    queryTables = worksheet.QueryTables;

                    for (int qt = 1; qt <= queryTables.Count; qt++)
                    {
                        dynamic? queryTable = null;
                        try
                        {
                            queryTable = queryTables.Item(qt);
                            string qtName = queryTable.Name?.ToString() ?? "";

                            if (qtName.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase) ||
                                qtName.Contains(queryName.Replace(" ", "_")))
                            {
                                string sheetName = worksheet.Name;
                                return sheetName;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTable);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref queryTables);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        catch { }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        return null;
    }

    /// <inheritdoc />
    public async Task<PowerQueryListResult> ListAsync(IExcelBatch batch)
    {
        var result = new PowerQueryListResult { FilePath = batch.WorkbookPath };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            try
            {
                queriesCollection = ctx.Book.Queries;
                int count = queriesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic? query = null;
                    try
                    {
                        query = queriesCollection.Item(i);
                        string name = query.Name ?? $"Query{i}";
                        string formula = query.Formula ?? "";

                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;

                        // Check if connection only
                        bool isConnectionOnly = true;
                        dynamic? connections = null;
                        try
                        {
                            connections = ctx.Book.Connections;
                            for (int c = 1; c <= connections.Count; c++)
                            {
                                dynamic? conn = null;
                                try
                                {
                                    conn = connections.Item(c);
                                    string connName = conn.Name?.ToString() ?? "";
                                    if (connName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                                        connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase))
                                    {
                                        isConnectionOnly = false;
                                        break;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref conn);
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            ComUtilities.Release(ref connections);
                        }

                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = name,
                            Formula = formula,
                            FormulaPreview = preview,
                            IsConnectionOnly = isConnectionOnly
                        });
                    }
                    catch (Exception queryEx)
                    {
                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = $"Error Query {i}",
                            Formula = "",
                            FormulaPreview = $"Error: {queryEx.Message}",
                            IsConnectionOnly = false
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref query);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error accessing Power Queries: {ex.Message}";

                string extension = Path.GetExtension(batch.WorkbookPath).ToLowerInvariant();
                if (extension == ".xls")
                {
                    result.ErrorMessage += " (.xls files don't support Power Query)";
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> ViewAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        return await batch.Execute<PowerQueryViewResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(ctx.Book);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return result;
                }

                string mCode = query.Formula;
                result.MCode = mCode;
                result.CharacterCount = mCode.Length;

                // Check if connection only
                bool isConnectionOnly = true;
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    for (int c = 1; c <= connections.Count; c++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(c);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                                connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                isConnectionOnly = false;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                catch { }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                result.IsConnectionOnly = isConnectionOnly;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error viewing query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateAsync(IExcelBatch batch, string queryName, string mCodeFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-update"
        };

        // Validate and normalize the M code file path to prevent path traversal attacks
        try
        {
            mCodeFile = PathValidator.ValidateExistingFile(mCodeFile, nameof(mCodeFile));
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid M code file path: {ex.Message}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        // STEP 1: Capture current load configuration BEFORE update
        var loadConfigBefore = await GetLoadConfigAsync(batch, queryName);

        // STEP 2: Update the query M code
        result = await batch.ExecuteAsync<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(ctx.Book);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return result;
                }

                // Update M code
                query.Formula = mCode;
                result.Success = true;

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.SuggestedNextActions =
                [
                    "Privacy levels cannot be set programmatically",
                    "Open file in Excel: File → Options → Privacy",
                    "Set appropriate privacy level or enable 'Ignore Privacy Levels' for testing",
                    "See documentation for privacy level guidance"
                ];
                result.WorkflowHint = "Privacy levels must be configured manually in Excel UI";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error updating query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });

        // STEP 3: Restore load configuration if query was loaded before
        if (result.Success && loadConfigBefore.Success)
        {
            if (loadConfigBefore.LoadMode == PowerQueryLoadMode.LoadToTable ||
                loadConfigBefore.LoadMode == PowerQueryLoadMode.LoadToBoth)
            {
                string targetSheet = loadConfigBefore.TargetSheet ?? queryName;
                var restoreResult = await SetLoadToTableAsync(batch, queryName, targetSheet);

                if (!restoreResult.Success)
                {
                    result.ErrorMessage = $"Query updated but failed to restore load configuration: {restoreResult.ErrorMessage}";
                    result.SuggestedNextActions =
                    [
                        "Query M code updated successfully",
                        "⚠️ Load configuration could not be restored automatically",
                        $"Manually load with: Use 'set-load-to-table' with worksheet '{targetSheet}'",
                        "Or use 'get-load-config' to check current state"
                    ];
                    return result;
                }

                // Successfully updated and restored load configuration
                result.SuggestedNextActions =
                [
                    "Query updated successfully, load configuration preserved",
                    "Data automatically refreshed with new M code",
                    "Use 'get-load-config' to verify configuration if needed"
                ];
                result.WorkflowHint = "Query updated successfully. M code changed, configuration preserved, data refreshed automatically.";
                return result;
            }
        }

        // Connection-only query or restore not needed
        if (result.Success)
        {
            result.SuggestedNextActions =
            [
                "Query updated successfully (connection-only)",
                "Use 'set-load-to-table' if you want to load data",
                "Use 'get-load-config' to verify configuration"
            ];
            result.WorkflowHint = "Query updated as connection-only (no data loaded).";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportAsync(IExcelBatch batch, string queryName, string outputFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-export"
        };

        // Validate and normalize the output file path to prevent path traversal attacks
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

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(ctx.Book);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return result;
                }

                string mCode = query.Formula;
                await File.WriteAllTextAsync(outputFile, mCode, ct);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> ImportAsync(IExcelBatch batch, string queryName, string mCodeFile, string loadDestination = "worksheet", string? worksheetName = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-import"
        };

        // Validate loadDestination parameter
        var validDestinations = new[] { "worksheet", "data-model", "both", "connection-only" };
        if (!validDestinations.Contains(loadDestination.ToLowerInvariant()))
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid loadDestination: '{loadDestination}'. Valid values: {string.Join(", ", validDestinations)}";
            return result;
        }

        // Validate and normalize the M code file path to prevent path traversal attacks
        try
        {
            mCodeFile = PathValidator.ValidateExistingFile(mCodeFile, nameof(mCodeFile));
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid M code file path: {ex.Message}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        result = await batch.ExecuteAsync<OperationResult>((ctx, ct) =>
        {
            dynamic? existingQuery = null;
            dynamic? queriesCollection = null;
            dynamic? newQuery = null;
            try
            {
                // Check if query already exists
                existingQuery = ComUtilities.FindQuery(ctx.Book, queryName);
                if (existingQuery != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' already exists. Use pq-update to modify it.";
                    return result;
                }

                // Add new query
                queriesCollection = ctx.Book.Queries;
                newQuery = queriesCollection.Add(queryName, mCode);

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.SuggestedNextActions =
                [
                    "Privacy levels cannot be set programmatically",
                    "Open file in Excel: File → Options → Privacy",
                    "Set appropriate privacy level or enable 'Ignore Privacy Levels' for testing",
                    "See documentation for privacy level guidance"
                ];
                result.WorkflowHint = "Privacy levels must be configured manually in Excel UI";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error importing query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref newQuery);
                ComUtilities.Release(ref queriesCollection);
                ComUtilities.Release(ref existingQuery);
            }
        });

        // Auto-load based on loadDestination parameter
        if (result.Success)
        {
            var destination = loadDestination.ToLowerInvariant();
            OperationResult? loadResult = null;

            switch (destination)
            {
                case "worksheet":
                    string targetSheet = worksheetName ?? queryName;
                    loadResult = await SetLoadToTableAsync(batch, queryName, targetSheet);
                    break;

                case "data-model":
                    var dmResult = await SetLoadToDataModelAsync(batch, queryName);
                    loadResult = new OperationResult
                    {
                        Success = dmResult.Success,
                        ErrorMessage = dmResult.ErrorMessage,
                        FilePath = dmResult.FilePath
                    };
                    break;

                case "both":
                    string targetSheetBoth = worksheetName ?? queryName;
                    var bothResult = await SetLoadToBothAsync(batch, queryName, targetSheetBoth);
                    loadResult = new OperationResult
                    {
                        Success = bothResult.Success,
                        ErrorMessage = bothResult.ErrorMessage,
                        FilePath = bothResult.FilePath
                    };
                    break;

                case "connection-only":
                    // No loading - query imported but not executed
                    result.SuggestedNextActions =
                    [
                        "Query imported as connection-only (NOT validated yet)",
                        "⚠️ M code has not been executed or validated",
                        "Use 'set-load-to-table' to validate and load data to worksheet",
                        "Use 'set-load-to-data-model' to load data to Power Pivot Data Model",
                        "Use 'view' to review imported M code"
                    ];
                    result.WorkflowHint = "Query imported as connection-only (M code not executed or validated).";
                    return result;
            }

            // Handle loading result
            if (loadResult != null && !loadResult.Success)
            {
                // Loading failed - query is imported but connection-only
                result.Success = true; // Import itself succeeded
                result.ErrorMessage = $"Query imported but failed to load to {destination}: {loadResult.ErrorMessage}";
                result.SuggestedNextActions =
                [
                    "Query imported as connection-only (auto-load failed)",
                    $"Try manually: Use appropriate set-load action",
                    "Or use 'view' to review M code for issues"
                ];
                result.WorkflowHint = $"Query imported but could not be automatically loaded to {destination}";
                return result;
            }
            else if (loadResult != null && loadResult.Success)
            {
                // CRITICAL: Save the workbook to persist changes
                await batch.SaveAsync();

                // Query was loaded successfully
                result.SuggestedNextActions =
                [
                    $"Query imported and data loaded to {destination} successfully",
                    "Use 'view' to inspect M code",
                    "Use 'get-load-config' to verify configuration"
                ];
                result.WorkflowHint = $"Query imported and loaded to {destination}.";
                return result;
            }
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryRefreshResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            RefreshTime = DateTime.Now
        };

        return await batch.Execute<PowerQueryRefreshResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(ctx.Book);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return result;
                }

                // Check if query has a connection to refresh
                dynamic? targetConnection = null;
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(i);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                                connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                targetConnection = conn;
                                conn = null; // Don't release - we're using it
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                catch { }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                if (targetConnection != null)
                {
                    try
                    {
                        // Attempt refresh and capture any errors
                        targetConnection.Refresh();

                        // Check for errors after refresh
                        result.HasErrors = false;
                        result.Success = true;
                        result.LoadedToSheet = DetermineLoadedSheet(ctx.Book, queryName);

                        // Add workflow guidance
                        result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterRefresh(
                            hasErrors: false,
                            isConnectionOnly: false);
                        result.WorkflowHint = PowerQueryWorkflowGuidance.GetWorkflowHint("pq-refresh", true);
                    }
                    catch (COMException comEx)
                    {
                        // Capture detailed error information
                        result.Success = false;
                        result.HasErrors = true;
                        result.ErrorMessages.Add(ParsePowerQueryError(comEx));
                        result.ErrorMessage = string.Join("; ", result.ErrorMessages);

                        var errorCategory = CategorizeError(comEx);
                        result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetErrorRecoverySteps(errorCategory);
                        result.WorkflowHint = PowerQueryWorkflowGuidance.GetWorkflowHint("pq-refresh", false);
                    }
                    finally
                    {
                        ComUtilities.Release(ref targetConnection);
                    }
                }
                else
                {
                    // Connection-only query
                    ComUtilities.Release(ref query);
                    result.Success = true;
                    result.IsConnectionOnly = true;
                    result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterRefresh(
                        hasErrors: false,
                        isConnectionOnly: true);
                    result.WorkflowHint = "Query is connection-only. Use set-load-to-table to load data.";
                }

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error refreshing query: {ex.Message}";
                result.SuggestedNextActions =
                [
                    "Unexpected error during refresh",
                    "Check that Excel file is not corrupted",
                    "Verify query exists and is accessible"
                ];
                return result;
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> ErrorsAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        return await batch.Execute<PowerQueryViewResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Try to get error information if available
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(i);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                                connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                // Connection found - query has been loaded
                                result.MCode = "No error information available through Excel COM interface";
                                result.Success = true;
                                return result;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                catch { }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                result.MCode = "Query is connection-only - no error information available";
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error checking query errors: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> LoadToAsync(IExcelBatch batch, string queryName, string sheetName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-loadto"
        };

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Find or create target sheet
                dynamic? sheets = null;
                dynamic? targetSheet = null;
                try
                {
                    sheets = ctx.Book.Worksheets;

                    for (int i = 1; i <= sheets.Count; i++)
                    {
                        dynamic? sheet = null;
                        try
                        {
                            sheet = sheets.Item(i);
                            if (sheet.Name == sheetName)
                            {
                                targetSheet = sheet;
                                sheet = null; // Don't release - we're using it
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }

                    if (targetSheet == null)
                    {
                        targetSheet = sheets.Add();
                        targetSheet.Name = sheetName;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sheets);
                }

                // Get the workbook connections to find our query
                dynamic? connections = null;
                dynamic? targetConnection = null;
                try
                {
                    connections = ctx.Book.Connections;

                    // Look for existing connection for this query
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(i);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                                connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                targetConnection = conn;
                                conn = null; // Don't release - we're using it
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                // If no connection exists, we need to create one by loading the query to table
                if (targetConnection == null)
                {
                    // Access the query through the Queries collection and load it to table
                    dynamic? queries = null;
                    dynamic? targetQuery = null;
                    dynamic? queryTables = null;
                    dynamic? queryTable = null;
                    dynamic? rangeObj = null;
                    try
                    {
                        queries = ctx.Book.Queries;

                        for (int i = 1; i <= queries.Count; i++)
                        {
                            dynamic? q = null;
                            try
                            {
                                q = queries.Item(i);
                                if (q.Name.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                                {
                                    targetQuery = q;
                                    q = null; // Don't release - we're using it
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref q);
                            }
                        }

                        if (targetQuery == null)
                        {
                            result.Success = false;
                            result.ErrorMessage = $"Query '{queryName}' not found in queries collection";
                            return result;
                        }

                        // Create a QueryTable using the Mashup provider
                        queryTables = targetSheet.QueryTables;
                        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                        string commandText = $"SELECT * FROM [{queryName}]";

                        rangeObj = targetSheet.Range["A1"];
                        queryTable = queryTables.Add(connectionString, rangeObj, commandText);
                        queryTable.Name = queryName.Replace(" ", "_");
                        queryTable.RefreshStyle = 1; // xlInsertDeleteCells

                        // Set additional properties for better data loading
                        queryTable.BackgroundQuery = false; // Don't run in background
                        queryTable.PreserveColumnInfo = true;
                        queryTable.PreserveFormatting = true;
                        queryTable.AdjustColumnWidth = true;

                        // Refresh to actually load the data
                        queryTable.Refresh(false); // false = wait for completion
                    }
                    finally
                    {
                        ComUtilities.Release(ref rangeObj);
                        ComUtilities.Release(ref queryTable);
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref targetQuery);
                        ComUtilities.Release(ref queries);
                    }
                }
                else
                {
                    // Connection exists, create QueryTable from existing connection
                    dynamic? queryTables = null;
                    dynamic? queryTable = null;
                    dynamic? rangeObj = null;
                    try
                    {
                        queryTables = targetSheet.QueryTables;

                        // Remove any existing QueryTable with the same name
                        try
                        {
                            for (int i = queryTables.Count; i >= 1; i--)
                            {
                                dynamic? qt = null;
                                try
                                {
                                    qt = queryTables.Item(i);
                                    if (qt.Name.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase))
                                    {
                                        qt.Delete();
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref qt);
                                }
                            }
                        }
                        catch { } // Ignore errors if no existing QueryTable

                        // Create new QueryTable
                        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                        string commandText = $"SELECT * FROM [{queryName}]";

                        rangeObj = targetSheet.Range["A1"];
                        queryTable = queryTables.Add(connectionString, rangeObj, commandText);
                        queryTable.Name = queryName.Replace(" ", "_");
                        queryTable.RefreshStyle = 1; // xlInsertDeleteCells
                        queryTable.BackgroundQuery = false;
                        queryTable.PreserveColumnInfo = true;
                        queryTable.PreserveFormatting = true;
                        queryTable.AdjustColumnWidth = true;

                        // Refresh to load data
                        queryTable.Refresh(false);
                    }
                    finally
                    {
                        ComUtilities.Release(ref rangeObj);
                        ComUtilities.Release(ref queryTable);
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref targetConnection);
                    }
                }

                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref query);
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error loading query to worksheet: {ex.Message}";
                return result;
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-delete"
        };

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? queriesCollection = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                queriesCollection = ctx.Book.Queries;
                queriesCollection.Item(queryName).Delete();

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <summary>
    /// Helper to get all query names
    /// </summary>
    private static List<string> GetQueryNames(dynamic workbook)
    {
        var names = new List<string>();
        dynamic? queriesCollection = null;
        try
        {
            queriesCollection = workbook.Queries;
            for (int i = 1; i <= queriesCollection.Count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = queriesCollection.Item(i);
                    names.Add(query.Name);
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }
        }
        catch { }
        finally
        {
            ComUtilities.Release(ref queriesCollection);
        }
        return names;
    }

    /// <inheritdoc />
    public async Task<WorksheetListResult> SourcesAsync(IExcelBatch batch)
    {
        var result = new WorksheetListResult { FilePath = batch.WorkbookPath };

        return await batch.Execute<WorksheetListResult>((ctx, ct) =>
        {
            dynamic? worksheets = null;
            dynamic? names = null;
            try
            {
                // Get all tables from all worksheets
                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? tables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        string wsName = worksheet.Name;

                        tables = worksheet.ListObjects;
                        for (int i = 1; i <= tables.Count; i++)
                        {
                            dynamic? table = null;
                            try
                            {
                                table = tables.Item(i);
                                result.Worksheets.Add(new WorksheetInfo
                                {
                                    Name = table.Name,
                                    Index = i,
                                    Visible = true
                                });
                            }
                            finally
                            {
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                // Get all named ranges
                names = ctx.Book.Names;
                int namedRangeIndex = result.Worksheets.Count + 1;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic? name = null;
                    try
                    {
                        name = names.Item(i);
                        string nameValue = name.Name;
                        if (!nameValue.StartsWith("_"))
                        {
                            result.Worksheets.Add(new WorksheetInfo
                            {
                                Name = nameValue,
                                Index = namedRangeIndex++,
                                Visible = true
                            });
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref name);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing sources: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref worksheets);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> TestAsync(IExcelBatch batch, string sourceName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-test"
        };

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            dynamic? tempQuery = null;
            try
            {
                // Create a test query to load the source
                string testQuery = $@"
let
    Source = Excel.CurrentWorkbook(){{[Name=""{sourceName.Replace("\"", "\"\"")}""]]}}[Content]
in
    Source";

                queriesCollection = ctx.Book.Queries;
                tempQuery = queriesCollection.Add("_TestQuery", testQuery);

                // Try to refresh
                bool refreshSuccess = false;
                try
                {
                    tempQuery.Refresh();
                    refreshSuccess = true;
                }
                catch { }

                // Clean up
                tempQuery.Delete();

                result.Success = true;
                if (!refreshSuccess)
                {
                    result.ErrorMessage = "Source exists but could not refresh (may need data source configuration)";
                }

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found or cannot be loaded: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref tempQuery);
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    /// <inheritdoc />
    public async Task<WorksheetDataResult> PeekAsync(IExcelBatch batch, string sourceName)
    {
        var result = new WorksheetDataResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sourceName
        };

        return await batch.Execute<WorksheetDataResult>((ctx, ct) =>
        {
            dynamic? names = null;
            dynamic? worksheets = null;
            try
            {
                // Check if it's a named range (single value)
                names = ctx.Book.Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic? name = null;
                    try
                    {
                        name = names.Item(i);
                        string nameValue = name.Name;
                        if (nameValue == sourceName)
                        {
                            try
                            {
                                var value = name.RefersToRange.Value;
                                result.Data.Add([value]);
                                result.RowCount = 1;
                                result.ColumnCount = 1;
                                result.Success = true;
                                return result;
                            }
                            catch
                            {
                                result.Success = false;
                                result.ErrorMessage = "Named range found but value cannot be read (may be #REF!)";
                                return result;
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref name);
                    }
                }

                // Check if it's a table
                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? tables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        tables = worksheet.ListObjects;
                        for (int i = 1; i <= tables.Count; i++)
                        {
                            dynamic? table = null;
                            dynamic? listCols = null;
                            try
                            {
                                table = tables.Item(i);
                                if (table.Name == sourceName)
                                {
                                    result.RowCount = table.ListRows.Count;
                                    result.ColumnCount = table.ListColumns.Count;

                                    // Get column names
                                    listCols = table.ListColumns;
                                    for (int c = 1; c <= Math.Min(result.ColumnCount, 10); c++)
                                    {
                                        dynamic? listCol = null;
                                        try
                                        {
                                            listCol = listCols.Item(c);
                                            result.Headers.Add(listCol.Name);
                                        }
                                        finally
                                        {
                                            ComUtilities.Release(ref listCol);
                                        }
                                    }

                                    result.Success = true;
                                    return result;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref listCols);
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error peeking source: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref names);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> EvalAsync(IExcelBatch batch, string mExpression)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = "_EvalExpression"
        };

        return await batch.Execute<PowerQueryViewResult>((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            dynamic? tempQuery = null;
            try
            {
                // Create a temporary query with the expression
                string evalQuery = $@"
let
    Result = {mExpression}
in
    Result";

                queriesCollection = ctx.Book.Queries;
                tempQuery = queriesCollection.Add("_EvalQuery", evalQuery);

                result.MCode = evalQuery;
                result.CharacterCount = evalQuery.Length;

                // Try to refresh
                try
                {
                    tempQuery.Refresh();
                    result.Success = true;
                    result.ErrorMessage = null;
                }
                catch (Exception refreshEx)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Expression syntax is valid but refresh failed: {refreshEx.Message}";
                }

                // Clean up
                tempQuery.Delete();

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Expression evaluation failed: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref tempQuery);
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetConnectionOnlyAsync(IExcelBatch batch, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-connection-only"
        };

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Remove any existing connections and QueryTables for this query
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting connection only: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadToTableResult> SetLoadToTableAsync(IExcelBatch batch, string queryName, string sheetName)
    {
        var result = new PowerQueryLoadToTableResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-table",
            QueryName = queryName,
            SheetName = sheetName,
            WorkflowStatus = "Failed"
        };

        return await batch.Execute<PowerQueryLoadToTableResult>((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? targetSheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;
            try
            {
                // STEP 1: Verify query exists
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    result.WorkflowStatus = "Failed";
                    result.SuggestedNextActions =
                    [
                        $"Use 'list' to see available queries",
                        $"Check the query name spelling: '{queryName}'"
                    ];
                    return result;
                }

                // STEP 2: Find or create target sheet
                sheets = ctx.Book.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        if (sheet.Name == sheetName)
                        {
                            targetSheet = sheet;
                            sheet = null; // Don't release - we're keeping it
                            break;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // STEP 3: Configure query (remove old connections, create new QueryTable)
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = queryName,
                    RefreshImmediately = true // CRITICAL: Refresh synchronously to persist QueryTable properly
                };
                PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, queryTableOptions);

                result.ConfigurationApplied = true;

                // Note: RefreshImmediately=true causes CreateQueryTable to call queryTable.Refresh(false)
                // which is SYNCHRONOUS and ensures proper persistence when workbook is saved.
                // This follows Microsoft's documented pattern: Create → Refresh(False) → Save
                // (See VBA example: https://learn.microsoft.com/en-us/office/troubleshoot/excel/...)
                // RefreshAll() is ASYNCHRONOUS and unreliable for individual QueryTable persistence.

                // STEP 4: VERIFY data was actually loaded
                queryTables = targetSheet.QueryTables;
                string normalizedName = queryName.Replace(" ", "_");
                bool foundQueryTable = false;
                int rowsLoaded = 0;

                for (int qt = 1; qt <= queryTables.Count; qt++)
                {
                    dynamic? qt_obj = null;
                    try
                    {
                        qt_obj = queryTables.Item(qt);
                        string qtName = qt_obj.Name?.ToString() ?? "";

                        if (qtName.Equals(normalizedName, StringComparison.OrdinalIgnoreCase) ||
                            qtName.Contains(normalizedName, StringComparison.OrdinalIgnoreCase))
                        {
                            foundQueryTable = true;

                            // Get row count from ResultRange
                            try
                            {
                                dynamic? resultRange = qt_obj.ResultRange;
                                if (resultRange != null)
                                {
                                    rowsLoaded = resultRange.Rows.Count;
                                    ComUtilities.Release(ref resultRange);
                                }
                            }
                            catch
                            {
                                // If we can't get row count, at least we found the QueryTable
                                rowsLoaded = 0;
                            }

                            queryTable = qt_obj;
                            qt_obj = null; // Keep reference
                            break;
                        }
                    }
                    finally
                    {
                        if (qt_obj != null)
                        {
                            ComUtilities.Release(ref qt_obj);
                        }
                    }
                }

                if (foundQueryTable)
                {
                    result.Success = true;
                    result.DataLoadedToTable = true;
                    result.RowsLoaded = rowsLoaded;
                    result.WorkflowStatus = "Complete";
                    result.WorkflowHint = $"Query '{queryName}' loaded to worksheet '{sheetName}' with {rowsLoaded} rows";
                    result.SuggestedNextActions =
                    [
                        $"View data in worksheet '{sheetName}'",
                        "Use 'refresh' to reload data from source",
                        "Create Excel tables or PivotTables from the data"
                    ];
                }
                else
                {
                    result.Success = false;
                    result.DataLoadedToTable = false;
                    result.RowsLoaded = 0;
                    result.WorkflowStatus = "Partial";
                    result.ErrorMessage = $"Configuration applied but QueryTable not found after refresh";
                    result.SuggestedNextActions =
                    [
                        "Check if query has valid data source",
                        "Verify privacy level settings",
                        "Use 'errors' action to see query errors"
                    ];
                }

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.WorkflowStatus = "Failed";
                result.WorkflowHint = "Privacy levels must be configured manually in Excel UI";
                result.SuggestedNextActions =
                [
                    "Privacy levels cannot be set programmatically",
                    "Open file in Excel: File → Options → Privacy",
                    "Set appropriate privacy level or enable 'Ignore Privacy Levels' for testing",
                    "See documentation for privacy level guidance"
                ];

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting load to table: {ex.Message}";
                result.WorkflowStatus = "Failed";
                result.SuggestedNextActions =
                [
                    "Check query name and worksheet name are valid",
                    "Verify Excel workbook is not corrupted",
                    "Review error message for specific issue"
                ];
                return result;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadToDataModelResult> SetLoadToDataModelAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadToDataModelResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-data-model",
            QueryName = queryName,
            ConfigurationApplied = false,
            DataLoadedToModel = false,
            RowsLoaded = 0,
            TablesInDataModel = 0,
            WorkflowStatus = "Failed"
        };

        return await batch.ExecuteAsync<PowerQueryLoadToDataModelResult>(async (ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // STEP 1: Configure query to load to data model
                // Remove existing table connections
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                // Configure Data Model loading using Connections.Add2
                bool configSuccess = SetQueryLoadToDataModel(ctx.Book, queryName, out string? configError);
                result.ConfigurationApplied = configSuccess;

                if (!configSuccess)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Failed to configure query for Data Model loading: {configError ?? "Unknown error"}";
                    result.WorkflowStatus = "Failed";
                    result.SuggestedNextActions =
                    [
                        "Check that the query exists using 'list'",
                        "Use 'view' to check query M code for syntax errors",
                        "Try 'set-load-to-table' first to validate query works",
                    ];
                    return result;
                }

                // STEP 2: Verify data was actually loaded to Data Model
                dynamic? model = null;
                dynamic? modelTables = null;
                try
                {
                    model = ctx.Book.Model;
                    modelTables = model.ModelTables;
                    result.TablesInDataModel = modelTables.Count;

                    // Find the query's table in the Data Model
                    bool foundTable = false;
                    int rowCount = 0;

                    for (int i = 1; i <= modelTables.Count; i++)
                    {
                        dynamic? table = null;
                        try
                        {
                            table = modelTables.Item(i);
                            string tableName = table.Name?.ToString() ?? "";

                            // Match by query name (Excel may add prefixes/suffixes)
                            if (tableName.Contains(queryName, StringComparison.OrdinalIgnoreCase))
                            {
                                foundTable = true;

                                // Get row count
                                try
                                {
                                    rowCount = (int)table.RecordCount;
                                }
                                catch
                                {
                                    rowCount = 0; // RecordCount may not be available immediately
                                }

                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref table);
                        }
                    }

                    result.DataLoadedToModel = foundTable;
                    result.RowsLoaded = rowCount;

                    if (foundTable)
                    {
                        result.Success = true;
                        result.WorkflowStatus = "Complete";
                        result.WorkflowHint = $"Power Query '{queryName}' loaded to Data Model with {rowCount} rows";
                        result.SuggestedNextActions =
                        [
                            "Create DAX measures using dm-create-measure",
                            "Add relationships using dm-create-relationship",
                            "View Data Model tables using dm-list-tables"
                        ];
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = "Query configured and refreshed, but table not found in Data Model";
                        result.WorkflowStatus = "Partial";
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                    ComUtilities.Release(ref model);
                }

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("protected", StringComparison.OrdinalIgnoreCase) || 
                                             comEx.Message.Contains("sensitivity label", StringComparison.OrdinalIgnoreCase))
            {
                // Microsoft Purview sensitivity label error - encrypted file
                // Get M code to extract file path from File.Contents()
                string? filePath = null;
                try
                {
                    var viewResult = await ViewAsync(batch, queryName);
                    if (viewResult.Success && !string.IsNullOrEmpty(viewResult.MCode))
                    {
                        filePath = ExtractFileContentsPath(viewResult.MCode);
                    }
                }
                catch (COMException)
                {
                    // If we can't get M code due to COM error, continue without file path
                }

                string filePathInfo = !string.IsNullOrEmpty(filePath) 
                    ? $"\n\nSource file: {filePath}" 
                    : "";

                result.Success = false;
                result.ErrorMessage = $"Source Excel file has Microsoft Purview sensitivity labels (encryption).{filePathInfo}\n\n" +
                                    "Power Query cannot access encrypted Excel files.\n\n" +
                                    "SOLUTION 1 (Recommended): Change sensitivity label to Public\n" +
                                    "  - Open the source file in Excel\n" +
                                    "  - Click Home tab → Sensitivity button → Select \"Public\" label\n" +
                                    "  - Save and close\n" +
                                    "  - Retry: excel_powerquery(action: 'set-load-to-data-model', queryName: '{queryName}')\n\n" +
                                    "SOLUTION 2: Modify M code to use different data source\n" +
                                    "  - Replace File.Contents() with Excel.CurrentWorkbook() if data is in same workbook\n" +
                                    "  - Export source data to CSV and use Csv.Document()\n" +
                                    "  - Use ODBC or SQL connection if source is a database\n\n" +
                                    "Technical details: https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations";
                
                result.WorkflowStatus = "Failed";
                result.WorkflowHint = "Microsoft Purview sensitivity labels prevent Power Query from reading encrypted Excel files";
                result.SuggestedNextActions =
                [
                    "Remove sensitivity label from source file (set to \"Public\")",
                    "OR modify M code to use non-encrypted data source",
                    "Power Query limitation: Cannot read files with sensitivity labels (other than Public/Non-Business)"
                ];

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.WorkflowStatus = "Failed";
                result.WorkflowHint = "Privacy levels must be configured manually in Excel UI";
                result.SuggestedNextActions =
                [
                    "Privacy levels cannot be set programmatically",
                    "Open file in Excel: File → Options → Privacy",
                    "Set appropriate privacy level or enable 'Ignore Privacy Levels' for testing",
                    "See documentation for privacy level guidance"
                ];

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error in atomic load-to-data-model operation: {ex.Message}";
                result.WorkflowStatus = "Failed";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadToBothResult> SetLoadToBothAsync(IExcelBatch batch, string queryName, string sheetName)
    {
        var result = new PowerQueryLoadToBothResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-both",
            QueryName = queryName,
            SheetName = sheetName,
            WorkflowStatus = "Failed"
        };

        return await batch.Execute<PowerQueryLoadToBothResult>((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? targetSheet = null;
            try
            {
                // STEP 1: Verify query exists
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    result.WorkflowStatus = "Failed";
                    return result;
                }

                // STEP 2: Find or create target sheet
                sheets = ctx.Book.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        if (sheet.Name == sheetName)
                        {
                            targetSheet = sheet;
                            sheet = null; // Don't release - we're keeping it
                            break;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // STEP 4: Configure query for BOTH table and Data Model loading
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                // Create QueryTable for worksheet loading
                var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = queryName,
                    RefreshImmediately = true // CRITICAL: Refresh synchronously to persist QueryTable properly
                };
                PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, queryTableOptions);

                // Configure query for Data Model loading
                if (!SetQueryLoadToDataModel(ctx.Book, queryName, out string? dmConfigError))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Failed to configure query for Data Model loading: {dmConfigError ?? "Unknown error"}";
                    result.WorkflowStatus = "Partial";
                    result.SuggestedNextActions =
                    [
                        "Use 'view' to check query M code for syntax errors",
                        "Try 'set-load-to-table' first to validate query works",
                        "Verify query data source is accessible and has data"
                    ];
                    return result;
                }

                result.ConfigurationApplied = true;

                // STEP 6: VERIFY data loaded to BOTH destinations
                bool foundInTable = false;
                bool foundInDataModel = false;
                int tableRows = 0;
                int modelRows = 0;
                int tablesInDataModel = 0;

                // Verify table loading
                dynamic? queryTables = null;
                try
                {
                    queryTables = targetSheet.QueryTables;
                    string normalizedName = queryName.Replace(" ", "_");

                    for (int qt = 1; qt <= queryTables.Count; qt++)
                    {
                        dynamic? qt_obj = null;
                        try
                        {
                            qt_obj = queryTables.Item(qt);
                            string qtName = qt_obj.Name?.ToString() ?? "";

                            if (qtName.Equals(normalizedName, StringComparison.OrdinalIgnoreCase) ||
                                qtName.Contains(normalizedName, StringComparison.OrdinalIgnoreCase))
                            {
                                foundInTable = true;

                                // Get row count from ResultRange
                                try
                                {
                                    dynamic? resultRange = qt_obj.ResultRange;
                                    if (resultRange != null)
                                    {
                                        tableRows = resultRange.Rows.Count;
                                        ComUtilities.Release(ref resultRange);
                                    }
                                }
                                catch
                                {
                                    tableRows = 0;
                                }
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref qt_obj);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref queryTables);
                }

                // Verify Data Model loading
                dynamic? model = null;
                dynamic? modelTables = null;
                try
                {
                    model = ctx.Book.Model;
                    if (model != null)
                    {
                        modelTables = model.ModelTables;
                        tablesInDataModel = modelTables.Count;

                        for (int t = 1; t <= modelTables.Count; t++)
                        {
                            dynamic? table = null;
                            try
                            {
                                table = modelTables.Item(t);
                                string tableName = table.Name?.ToString() ?? "";

                                if (tableName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                                {
                                    foundInDataModel = true;
                                    modelRows = table.RecordCount;
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                    ComUtilities.Release(ref model);
                }

                // Set result based on verification
                result.DataLoadedToTable = foundInTable;
                result.DataLoadedToModel = foundInDataModel;
                result.RowsLoadedToTable = tableRows;
                result.RowsLoadedToModel = modelRows;
                result.TablesInDataModel = tablesInDataModel;

                if (foundInTable && foundInDataModel)
                {
                    result.Success = true;
                    result.WorkflowStatus = "Complete";
                    result.WorkflowHint = $"Query '{queryName}' loaded to both worksheet '{sheetName}' ({tableRows} rows) and Data Model ({modelRows} rows)";
                    result.SuggestedNextActions =
                    [
                        $"View data in worksheet '{sheetName}'",
                        "Create PivotTables using Data Model",
                        "Use 'refresh' to reload data from source"
                    ];
                }
                else if (foundInTable && !foundInDataModel)
                {
                    result.Success = false;
                    result.WorkflowStatus = "Partial";
                    result.ErrorMessage = "Data loaded to table but not to Data Model";
                    result.SuggestedNextActions =
                    [
                        "Verify query refreshed successfully",
                        "Check query configuration with 'get-load-config'",
                        "Try refreshing again or reload to Data Model only"
                    ];
                }
                else if (!foundInTable && foundInDataModel)
                {
                    result.Success = false;
                    result.WorkflowStatus = "Partial";
                    result.ErrorMessage = "Data loaded to Data Model but not to table";
                    result.SuggestedNextActions =
                    [
                        "Check worksheet and QueryTable configuration",
                        "Verify target sheet exists",
                        "Try 'set-load-to-both' again to load to both destinations"
                    ];
                }
                else
                {
                    result.Success = false;
                    result.WorkflowStatus = "Failed";
                    result.ErrorMessage = "Data not loaded to either destination";
                    result.SuggestedNextActions =
                    [
                        "Check query syntax and data source",
                        "Review privacy level settings",
                        "Use 'errors' action to see query errors"
                    ];
                }

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.WorkflowStatus = "Failed";
                result.WorkflowHint = "Privacy levels must be configured manually in Excel UI";
                result.SuggestedNextActions =
                [
                    "Privacy levels cannot be set programmatically",
                    "Open file in Excel: File → Options → Privacy",
                    "Set appropriate privacy level or enable 'Ignore Privacy Levels' for testing",
                    "See documentation for privacy level guidance"
                ];

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error in atomic load-to-both operation: {ex.Message}";
                result.WorkflowStatus = "Failed";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        return await batch.Execute<PowerQueryLoadConfigResult>((ctx, ct) =>
        {
            try
            {
                dynamic query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Check for QueryTables first (table loading)
                bool hasTableConnection = false;
                bool hasDataModelConnection = false;
                string? targetSheet = null;

                dynamic worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    dynamic queryTables = worksheet.QueryTables;

                    for (int qt = 1; qt <= queryTables.Count; qt++)
                    {
                        try
                        {
                            dynamic queryTable = queryTables.Item(qt);
                            string qtName = queryTable.Name?.ToString() ?? "";

                            // Check if this QueryTable is for our query
                            if (qtName.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase) ||
                                qtName.Contains(queryName.Replace(" ", "_")))
                            {
                                hasTableConnection = true;
                                targetSheet = worksheet.Name;
                                break;
                            }
                        }
                        catch
                        {
                            // Skip invalid QueryTables
                            continue;
                        }
                    }
                    if (hasTableConnection) break;
                }

                // Check for connections (for data model or other types)
                dynamic connections = ctx.Book.Connections;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic conn = connections.Item(i);
                    string connName = conn.Name?.ToString() ?? "";

                    if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        result.HasConnection = true;

                        // If we don't have a table connection but have a workbook connection,
                        // it's likely a data model connection
                        if (!hasTableConnection)
                        {
                            hasDataModelConnection = true;
                        }
                    }
                    else if (connName.Equals($"DataModel_{queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        // This is our explicit data model connection marker
                        result.HasConnection = true;
                        hasDataModelConnection = true;
                    }
                }

                // Always check for named range markers that indicate data model loading
                // (even if we have table connections, for LoadToBoth mode)
                if (!hasDataModelConnection)
                {
                    // Check for our data model marker
                    try
                    {
                        dynamic names = ctx.Book.Names;
                        string markerName = $"DataModel_Query_{queryName}";

                        for (int i = 1; i <= names.Count; i++)
                        {
                            try
                            {
                                dynamic existingName = names.Item(i);
                                if (existingName.Name.ToString() == markerName)
                                {
                                    hasDataModelConnection = true;
                                    break;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        // Cannot check names
                    }

                    // Fallback: Check if the query has data model indicators
                    if (!hasDataModelConnection)
                    {
                        hasDataModelConnection = CheckQueryDataModelConfiguration(query, ctx.Book);
                    }
                }

                // Determine load mode
                if (hasTableConnection && hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToBoth;
                }
                else if (hasTableConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToTable;
                }
                else if (hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToDataModel;
                }
                else
                {
                    result.LoadMode = PowerQueryLoadMode.ConnectionOnly;
                }

                result.TargetSheet = targetSheet;
                result.IsLoadedToDataModel = hasDataModelConnection;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error getting load config: {ex.Message}";
                return result;
            }
        });
    }

    /// <summary>
    /// Helper method to remove existing query connections and QueryTables
    /// </summary>
    private static void RemoveQueryConnections(dynamic workbook, string queryName)
    {
        dynamic? connections = null;
        dynamic? worksheets = null;
        try
        {
            // Remove connections
            connections = workbook.Connections;
            for (int i = connections.Count; i >= 1; i--)
            {
                dynamic? conn = null;
                try
                {
                    conn = connections.Item(i);
                    string connName = conn.Name?.ToString() ?? "";
                    if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        conn.Delete();
                    }
                }
                finally
                {
                    ComUtilities.Release(ref conn);
                }
            }

            // Remove QueryTables
            worksheets = workbook.Worksheets;
            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic? worksheet = null;
                dynamic? queryTables = null;
                try
                {
                    worksheet = worksheets.Item(ws);
                    queryTables = worksheet.QueryTables;

                    for (int qt = queryTables.Count; qt >= 1; qt--)
                    {
                        dynamic? queryTable = null;
                        try
                        {
                            queryTable = queryTables.Item(qt);
                            if (queryTable.Name?.ToString()?.Contains(queryName.Replace(" ", "_")) == true)
                            {
                                queryTable.Delete();
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTable);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref queryTables);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        catch
        {
            // Ignore errors when removing connections
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
            ComUtilities.Release(ref connections);
        }
    }

    /// <summary>
    /// Helper method to create a QueryTable connection that loads data to worksheet
    /// </summary>
    private static void CreateQueryTableConnection(dynamic workbook, dynamic targetSheet, string queryName)
    {
        try
        {
            // Ensure the query exists and is accessible
            dynamic query = ComUtilities.FindQuery(workbook, queryName);
            if (query == null)
            {
                throw new InvalidOperationException($"Query '{queryName}' not found");
            }

            // Get the QueryTables collection
            dynamic queryTables = targetSheet.QueryTables;

            // Build connection string for Power Query
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"SELECT * FROM [{queryName}]";

            // Get the target range - ensure it's valid
            dynamic startRange = targetSheet.Range["A1"];

            // Create the QueryTable
            dynamic queryTable = queryTables.Add(connectionString, startRange, commandText);
            queryTable.Name = queryName.Replace(" ", "_");
            queryTable.RefreshStyle = 1; // xlInsertDeleteCells
            queryTable.BackgroundQuery = false;
            queryTable.PreserveColumnInfo = true;
            queryTable.PreserveFormatting = true;
            queryTable.AdjustColumnWidth = true;
            queryTable.RefreshOnFileOpen = false;
            queryTable.SavePassword = false;

            // Refresh to load data immediately
            queryTable.Refresh(false);
        }
        catch (COMException comEx)
        {
            // Provide more detailed error information
            string hexCode = $"0x{comEx.HResult:X}";
            throw new InvalidOperationException(
                $"Failed to create QueryTable connection for '{queryName}': {comEx.Message} (Error: {hexCode}). " +
                $"This may occur if the query needs to be refreshed first or if there are data source connectivity issues.",
                comEx);
        }
    }

    /// <summary>
    /// Configures a Power Query to load to Data Model using Excel COM API
    /// Based on validated VBA pattern using Connections.Add2 method
    /// Reference: Working VBA code that successfully loads queries to Data Model
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="queryName">Name of the query to configure</param>
    /// <param name="errorMessage">Detailed error message if configuration fails</param>
    /// <returns>True if configuration succeeded, false if exception caught</returns>
    private static bool SetQueryLoadToDataModel(dynamic workbook, string queryName, out string? errorMessage)
    {
        dynamic? connections = null;
        dynamic? newConnection = null;
        errorMessage = null;

        try
        {
            connections = workbook.Connections;

            // Remove existing connections for this query to avoid conflicts
            ConnectionHelpers.RemoveConnections(workbook, queryName);

            // Use Connections.Add2 method (Excel 2013+) with Data Model parameters
            // This is the Microsoft-documented approach for loading Power Query to Data Model
            // Based on working VBA pattern:
            // w.Connections.Add2 "Query - " & query.Name, _
            //     "Connection to the '" & query.Name & "' query in the workbook.", _
            //     "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name, _
            //     """" & query.Name & """", 6, True, False

            string connectionName = $"Query - {queryName}";
            string description = $"Connection to the '{queryName}' query in the workbook.";
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"\"{queryName}\""; // Quoted query name
            int commandType = 6; // Data Model command type (xlCmdDAX or similar)
            bool createModelConnection = true; // CRITICAL: This loads to Data Model
            bool importRelationships = false;

            newConnection = connections.Add2(
                connectionName,
                description,
                connectionString,
                commandText,
                commandType,
                createModelConnection,
                importRelationships
            );

            return true;
        }
        catch (Exception ex)
        {
            // Capture detailed error for user feedback
            errorMessage = ex.Message;
            System.Diagnostics.Debug.WriteLine($"Failed to configure Data Model loading: {ex.Message}");
            return false;
        }
        finally
        {
            ComUtilities.Release(ref newConnection);
            ComUtilities.Release(ref connections);
        }
    }

    /// <summary>
    /// Check if a query is configured for data model loading
    /// </summary>
    private static bool CheckQueryDataModelConfiguration(dynamic query, dynamic workbook)
    {
        try
        {
            // Method 1: Check if the query has LoadToWorksheetModel property set
            try
            {
                bool loadToModel = query.LoadToWorksheetModel;
                if (loadToModel) return true;
            }
            catch
            {
                // Property doesn't exist
            }

            // Method 2: Check if query has ModelConnection property
            try
            {
                dynamic modelConnection = query.ModelConnection;
                if (modelConnection != null) return true;
            }
            catch
            {
                // Property doesn't exist
            }

            // Since we now use explicit DataModel_ connection markers,
            // this method is mainly for detecting native Excel data model configurations
            return false;
        }
        catch
        {
            return false;
        }
    }
}

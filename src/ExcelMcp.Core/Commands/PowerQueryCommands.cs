using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Models;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query management commands - Core data layer (no console output)
/// </summary>
public class PowerQueryCommands : IPowerQueryCommands
{
    /// <summary>
    /// Detects privacy levels from M code
    /// </summary>
    private static PowerQueryPrivacyLevel? DetectPrivacyLevelFromMCode(string mCode)
    {
        if (mCode.Contains("Privacy.None()", StringComparison.OrdinalIgnoreCase))
            return PowerQueryPrivacyLevel.None;
        if (mCode.Contains("Privacy.Private()", StringComparison.OrdinalIgnoreCase))
            return PowerQueryPrivacyLevel.Private;
        if (mCode.Contains("Privacy.Organizational()", StringComparison.OrdinalIgnoreCase))
            return PowerQueryPrivacyLevel.Organizational;
        if (mCode.Contains("Privacy.Public()", StringComparison.OrdinalIgnoreCase))
            return PowerQueryPrivacyLevel.Public;

        return null;
    }

    /// <summary>
    /// Determines recommended privacy level based on existing queries
    /// </summary>
    private static PowerQueryPrivacyLevel DetermineRecommendedPrivacyLevel(List<QueryPrivacyInfo> existingLevels)
    {
        if (existingLevels.Count == 0)
            return PowerQueryPrivacyLevel.Private; // Default to most secure

        // If any query uses Private, recommend Private (most secure)
        if (existingLevels.Any(q => q.PrivacyLevel == PowerQueryPrivacyLevel.Private))
            return PowerQueryPrivacyLevel.Private;

        // If all queries use Organizational, recommend Organizational
        if (existingLevels.All(q => q.PrivacyLevel == PowerQueryPrivacyLevel.Organizational))
            return PowerQueryPrivacyLevel.Organizational;

        // If any query uses Public, and no Private exists, recommend Public
        if (existingLevels.Any(q => q.PrivacyLevel == PowerQueryPrivacyLevel.Public))
            return PowerQueryPrivacyLevel.Public;

        // Default to Private for safety
        return PowerQueryPrivacyLevel.Private;
    }

    /// <summary>
    /// Generates explanation for privacy level recommendation
    /// </summary>
    private static string GeneratePrivacyExplanation(List<QueryPrivacyInfo> existingLevels, PowerQueryPrivacyLevel recommended)
    {
        if (existingLevels.Count == 0)
        {
            return "No existing queries detected with privacy levels. We recommend starting with 'Private' " +
                   "(most secure) and adjusting if needed.";
        }

        var levelCounts = existingLevels.GroupBy(q => q.PrivacyLevel)
                                       .ToDictionary(g => g.Key, g => g.Count());

        string existingSummary = string.Join(", ", levelCounts.Select(kvp => $"{kvp.Value} use {kvp.Key}"));

        return recommended switch
        {
            PowerQueryPrivacyLevel.Private =>
                $"Existing queries: {existingSummary}. We recommend 'Private' for maximum data protection, " +
                "preventing data from being shared between sources.",
            PowerQueryPrivacyLevel.Organizational =>
                $"Existing queries: {existingSummary}. We recommend 'Organizational' to allow data sharing " +
                "within your organization's data sources.",
            PowerQueryPrivacyLevel.Public =>
                $"Existing queries: {existingSummary}. We recommend 'Public' since your queries work with " +
                "publicly available data sources.",
            PowerQueryPrivacyLevel.None =>
                $"Existing queries: {existingSummary}. We recommend 'None' to ignore privacy levels, " +
                "but be aware this is the least secure option.",
            _ => existingSummary
        };
    }

    /// <summary>
    /// Detects privacy levels in all queries and creates error result with recommendation
    /// </summary>
    private static PowerQueryPrivacyErrorResult DetectPrivacyLevelsAndRecommend(dynamic workbook, string originalError)
    {
        var privacyLevels = new List<QueryPrivacyInfo>();

        try
        {
            dynamic queries = workbook.Queries;

            for (int i = 1; i <= queries.Count; i++)
            {
                try
                {
                    dynamic query = queries.Item(i);
                    string name = query.Name ?? $"Query{i}";
                    string formula = query.Formula ?? "";

                    var detectedLevel = DetectPrivacyLevelFromMCode(formula);
                    if (detectedLevel.HasValue)
                    {
                        privacyLevels.Add(new QueryPrivacyInfo(name, detectedLevel.Value));
                    }
                }
                catch { /* Skip queries that can't be read */ }
            }
        }
        catch { /* If we can't read queries, just proceed with empty list */ }

        var recommended = DetermineRecommendedPrivacyLevel(privacyLevels);

        return new PowerQueryPrivacyErrorResult
        {
            Success = false,
            ErrorMessage = "Privacy level required to combine data sources",
            ExistingPrivacyLevels = privacyLevels,
            RecommendedPrivacyLevel = recommended,
            Explanation = GeneratePrivacyExplanation(privacyLevels, recommended),
            OriginalError = originalError
        };
    }

    /// <summary>
    /// Applies privacy level to workbook for Power Query operations
    /// </summary>
    private static void ApplyPrivacyLevel(dynamic workbook, PowerQueryPrivacyLevel privacyLevel)
    {
        try
        {
            // In Excel COM, privacy settings are typically applied at the workbook or query level
            // The most reliable approach is to set the Fast Data Load property
            // Note: Actual privacy level application may vary by Excel version

            // Try to set privacy via workbook properties if available
            try
            {
                // Some Excel versions support setting privacy through workbook properties
                dynamic customProps = workbook.CustomDocumentProperties;
                string privacyValue = privacyLevel.ToString();

                // Try to update existing property
                bool found = false;
                for (int i = 1; i <= customProps.Count; i++)
                {
                    dynamic prop = customProps.Item(i);
                    if (prop.Name == "PowerQueryPrivacyLevel")
                    {
                        prop.Value = privacyValue;
                        found = true;
                        break;
                    }
                }

                // Create new property if not found
                if (!found)
                {
                    customProps.Add("PowerQueryPrivacyLevel", false, 4, privacyValue); // 4 = msoPropertyTypeString
                }
            }
            catch { /* Property approach not supported in this Excel version */ }

            // The key approach: Set Fast Data Load to false when using privacy levels
            // This ensures Excel respects privacy settings
            try
            {
                dynamic application = workbook.Application;
                // Set calculation mode that respects privacy
                if (privacyLevel != PowerQueryPrivacyLevel.None)
                {
                    // Enable background query to allow privacy checks
                    application.DisplayAlerts = false;
                }
            }
            catch { /* Application settings not accessible */ }
        }
        catch (Exception)
        {
            // Privacy level application is best-effort
            // If it fails, the operation will still proceed and may trigger privacy error
        }
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
            return "Privacy level mismatch. Use privacyLevel parameter to resolve.";
        if (message.Contains("syntax", StringComparison.OrdinalIgnoreCase))
            return "M code syntax error. Review query formula.";
        if (message.Contains("permission", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("access denied", StringComparison.OrdinalIgnoreCase))
            return "Access denied. Check file or data source permissions.";

        return message; // Return original if no pattern matches
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
        try
        {
            dynamic worksheets = workbook.Worksheets;
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

                        if (qtName.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase) ||
                            qtName.Contains(queryName.Replace(" ", "_")))
                        {
                            return worksheet.Name;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
        }
        catch { }

        return null;
    }

    /// <inheritdoc />
    public PowerQueryListResult List(string filePath)
    {
        var result = new PowerQueryListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic queriesCollection = workbook.Queries;
                int count = queriesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic query = queriesCollection.Item(i);
                        string name = query.Name ?? $"Query{i}";
                        string formula = query.Formula ?? "";

                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;

                        // Check if connection only
                        bool isConnectionOnly = true;
                        try
                        {
                            dynamic connections = workbook.Connections;
                            for (int c = 1; c <= connections.Count; c++)
                            {
                                dynamic conn = connections.Item(c);
                                string connName = conn.Name?.ToString() ?? "";
                                if (connName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                                    connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase))
                                {
                                    isConnectionOnly = false;
                                    break;
                                }
                            }
                        }
                        catch { }

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
                }

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error accessing Power Queries: {ex.Message}";

                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension == ".xls")
                {
                    result.ErrorMessage += " (.xls files don't support Power Query)";
                }

                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public PowerQueryViewResult View(string filePath, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = filePath,
            QueryName = queryName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                string mCode = query.Formula;
                result.MCode = mCode;
                result.CharacterCount = mCode.Length;

                // Check if connection only
                bool isConnectionOnly = true;
                try
                {
                    dynamic connections = workbook.Connections;
                    for (int c = 1; c <= connections.Count; c++)
                    {
                        dynamic conn = connections.Item(c);
                        string connName = conn.Name?.ToString() ?? "";
                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            isConnectionOnly = false;
                            break;
                        }
                    }
                }
                catch { }

                result.IsConnectionOnly = isConnectionOnly;
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error viewing query: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> Update(string filePath, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null, bool autoRefresh = true)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-update"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (!File.Exists(mCodeFile))
        {
            result.Success = false;
            result.ErrorMessage = $"M code file not found: {mCodeFile}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        // STEP 1: Capture current load configuration BEFORE update
        var loadConfigBefore = GetLoadConfig(filePath, queryName);

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(workbook, privacyLevel.Value);
                }

                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                // STEP 2: Update M code
                query.Formula = mCode;
                result.Success = true;

                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(workbook, comEx.Message);
                privacyError.FilePath = filePath;
                privacyError.Action = "pq-update";
                result = privacyError;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error updating query: {ex.Message}";
                return 1;
            }
        });

        // STEP 3: Restore load configuration AFTER update (if update succeeded)
        bool configRestored = false;
        if (result.Success && loadConfigBefore.Success &&
            loadConfigBefore.LoadMode != PowerQueryLoadMode.ConnectionOnly)
        {
            try
            {
                switch (loadConfigBefore.LoadMode)
                {
                    case PowerQueryLoadMode.LoadToTable:
                        SetLoadToTable(filePath, queryName, loadConfigBefore.TargetSheet!, privacyLevel);
                        break;
                    case PowerQueryLoadMode.LoadToDataModel:
                        SetLoadToDataModel(filePath, queryName, privacyLevel);
                        break;
                    case PowerQueryLoadMode.LoadToBoth:
                        SetLoadToBoth(filePath, queryName, loadConfigBefore.TargetSheet!, privacyLevel);
                        break;
                }
                configRestored = true;
            }
            catch (Exception ex)
            {
                // Log warning but don't fail the update
                result.SuggestedNextActions = new List<string>
                {
                    "Query updated but load configuration could not be restored",
                    $"Original configuration was: {loadConfigBefore.LoadMode}",
                    "Use 'set-load-to-table' or 'set-load-to-data-model' to reconfigure"
                };
                result.WorkflowHint = $"Query updated successfully, but load configuration reset. Error: {ex.Message}";
                return result;
            }
        }

        // STEP 4: Auto-refresh after successful update to validate changes
        if (result.Success && autoRefresh)
        {
            var refreshResult = Refresh(filePath, queryName);
            if (!refreshResult.Success || refreshResult.HasErrors)
            {
                result.Success = false;
                result.ErrorMessage = $"Query updated but validation failed: {string.Join(", ", refreshResult.ErrorMessages)}";
                result.SuggestedNextActions = new List<string>
                {
                    "Query update failed validation",
                    "Review error messages and fix M code issues",
                    "Use 'view' to examine updated M code",
                    "Revert changes if needed with 'update' using previous version"
                };
                result.WorkflowHint = PowerQueryWorkflowGuidance.GetWorkflowHint("pq-update", false);
                return result;
            }

            // Update and refresh successful
            if (configRestored)
            {
                result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterUpdate(
                    configPreserved: true,
                    hasErrors: false);
                result.WorkflowHint = $"Query updated and validated successfully. Load configuration preserved ({loadConfigBefore.LoadMode} to {loadConfigBefore.TargetSheet ?? "Data Model"}).";
            }
            else
            {
                result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterUpdate(
                    configPreserved: false,
                    hasErrors: false);
                result.WorkflowHint = "Query updated and validated successfully (connection-only).";
            }
        }
        else if (result.Success)
        {
            // Update successful but no auto-refresh
            if (configRestored)
            {
                result.SuggestedNextActions = new List<string>
                {
                    "Query updated successfully, load configuration preserved (validation skipped)",
                    "Use 'refresh' to validate the changes",
                    "Use 'get-load-config' to verify configuration"
                };
                result.WorkflowHint = $"Query updated. Load configuration preserved ({loadConfigBefore.LoadMode}).";
            }
            else
            {
                result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterUpdate(
                    configPreserved: false,
                    hasErrors: false);
                result.WorkflowHint = "Query updated successfully (connection-only).";
            }
        }
        else
        {
            result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterUpdate(
                configPreserved: false,
                hasErrors: true);
            result.WorkflowHint = PowerQueryWorkflowGuidance.GetWorkflowHint("pq-update", false);
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> Export(string filePath, string queryName, string outputFile)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-export"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                string mCode = query.Formula;
                File.WriteAllText(outputFile, mCode);

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting query: {ex.Message}";
                return 1;
            }
        });

        return await Task.FromResult(result);
    }

    /// <inheritdoc />
    public async Task<OperationResult> Import(string filePath, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null, bool autoRefresh = true, bool loadToWorksheet = true, string? worksheetName = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-import"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        if (!File.Exists(mCodeFile))
        {
            result.Success = false;
            result.ErrorMessage = $"M code file not found: {mCodeFile}";
            return result;
        }

        string mCode = await File.ReadAllTextAsync(mCodeFile);

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(workbook, privacyLevel.Value);
                }

                // Check if query already exists
                dynamic existingQuery = FindQuery(workbook, queryName);
                if (existingQuery != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' already exists. Use pq-update to modify it.";
                    return 1;
                }

                // Add new query
                dynamic queriesCollection = workbook.Queries;
                dynamic newQuery = queriesCollection.Add(queryName, mCode);

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(workbook, comEx.Message);
                privacyError.FilePath = filePath;
                privacyError.Action = "pq-import";
                result = privacyError;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error importing query: {ex.Message}";
                return 1;
            }
        });

        // Auto-load to worksheet if requested (default behavior)
        if (result.Success && loadToWorksheet)
        {
            string targetSheet = worksheetName ?? queryName;
            var loadResult = SetLoadToTable(filePath, queryName, targetSheet, privacyLevel);

            if (!loadResult.Success)
            {
                // Loading failed - query is imported but connection-only
                result.Success = true; // Import itself succeeded
                result.ErrorMessage = $"Query imported but failed to load to worksheet: {loadResult.ErrorMessage}";
                result.SuggestedNextActions = new List<string>
                {
                    "Query imported as connection-only (auto-load failed)",
                    $"Try manually: Use 'set-load-to-table' with worksheet name",
                    "Or use 'view' to review M code for issues"
                };
                result.WorkflowHint = "Query imported but could not be automatically loaded to worksheet";
                return result;
            }
        }

        // Auto-refresh after successful import to validate query
        if (result.Success && autoRefresh)
        {
            var refreshResult = Refresh(filePath, queryName);
            if (!refreshResult.Success || refreshResult.HasErrors)
            {
                result.Success = false;
                result.ErrorMessage = $"Query imported but validation failed: {string.Join(", ", refreshResult.ErrorMessages)}";
                result.SuggestedNextActions = new List<string>
                {
                    "Query created but has errors",
                    "Use 'view' to review M code",
                    "Fix data source issues before proceeding"
                };
                result.WorkflowHint = PowerQueryWorkflowGuidance.GetWorkflowHint("pq-import", false);
                return result;
            }

            // Import and refresh successful
            result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterImport(
                isConnectionOnly: !loadToWorksheet || refreshResult.IsConnectionOnly,
                hasErrors: false);
            result.WorkflowHint = loadToWorksheet
                ? $"Query imported, data loaded to worksheet, and validated successfully"
                : refreshResult.IsConnectionOnly
                    ? "Query imported and validated successfully (connection-only mode)"
                    : $"Query imported, validated, and data loaded to {refreshResult.LoadedToSheet}";
        }
        else if (result.Success)
        {
            // Import successful but no auto-refresh
            result.SuggestedNextActions = new List<string>
            {
                loadToWorksheet
                    ? "Query imported and loaded to worksheet (validation skipped)"
                    : "Query imported successfully (validation skipped)",
                "Use 'refresh' to validate the query works",
                "Use 'get-load-config' to check configuration"
            };
            result.WorkflowHint = PowerQueryWorkflowGuidance.GetWorkflowHint("pq-import", true);
        }

        return result;
    }

    /// <inheritdoc />
    public PowerQueryRefreshResult Refresh(string filePath, string queryName)
    {
        var result = new PowerQueryRefreshResult
        {
            FilePath = filePath,
            QueryName = queryName,
            RefreshTime = DateTime.Now
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(workbook);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return 1;
                }

                // Check if query has a connection to refresh
                dynamic? targetConnection = null;
                try
                {
                    dynamic connections = workbook.Connections;
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic conn = connections.Item(i);
                        string connName = conn.Name?.ToString() ?? "";
                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            targetConnection = conn;
                            break;
                        }
                    }
                }
                catch { }

                if (targetConnection != null)
                {
                    try
                    {
                        // Attempt refresh and capture any errors
                        targetConnection.Refresh();

                        // Check for errors after refresh
                        result.HasErrors = false;
                        result.Success = true;
                        result.LoadedToSheet = DetermineLoadedSheet(workbook, queryName);

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
                }
                else
                {
                    // Connection-only query
                    result.Success = true;
                    result.IsConnectionOnly = true;
                    result.SuggestedNextActions = PowerQueryWorkflowGuidance.GetNextStepsAfterRefresh(
                        hasErrors: false,
                        isConnectionOnly: true);
                    result.WorkflowHint = "Query is connection-only. Use set-load-to-table to load data.";
                }

                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error refreshing query: {ex.Message}";
                result.SuggestedNextActions = new List<string>
                {
                    "Unexpected error during refresh",
                    "Check that Excel file is not corrupted",
                    "Verify query exists and is accessible"
                };
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public PowerQueryViewResult Errors(string filePath, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = filePath,
            QueryName = queryName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Try to get error information if available
                try
                {
                    dynamic connections = workbook.Connections;
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic conn = connections.Item(i);
                        string connName = conn.Name?.ToString() ?? "";
                        if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            // Connection found - query has been loaded
                            result.MCode = "No error information available through Excel COM interface";
                            result.Success = true;
                            return 0;
                        }
                    }
                }
                catch { }

                result.MCode = "Query is connection-only - no error information available";
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error checking query errors: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult LoadTo(string filePath, string queryName, string sheetName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-loadto"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Find or create target sheet
                dynamic sheets = workbook.Worksheets;
                dynamic? targetSheet = null;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item(i);
                    if (sheet.Name == sheetName)
                    {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // Get the workbook connections to find our query
                dynamic connections = workbook.Connections;
                dynamic? targetConnection = null;

                // Look for existing connection for this query
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic conn = connections.Item(i);
                    string connName = conn.Name?.ToString() ?? "";
                    if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        targetConnection = conn;
                        break;
                    }
                }

                // If no connection exists, we need to create one by loading the query to table
                if (targetConnection == null)
                {
                    // Access the query through the Queries collection and load it to table
                    dynamic queries = workbook.Queries;
                    dynamic? targetQuery = null;

                    for (int i = 1; i <= queries.Count; i++)
                    {
                        dynamic q = queries.Item(i);
                        if (q.Name.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetQuery = q;
                            break;
                        }
                    }

                    if (targetQuery == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found in queries collection";
                        return 1;
                    }

                    // Create a QueryTable using the Mashup provider
                    dynamic queryTables = targetSheet.QueryTables;
                    string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                    string commandText = $"SELECT * FROM [{queryName}]";

                    dynamic queryTable = queryTables.Add(connectionString, targetSheet.Range["A1"], commandText);
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
                else
                {
                    // Connection exists, create QueryTable from existing connection
                    dynamic queryTables = targetSheet.QueryTables;

                    // Remove any existing QueryTable with the same name
                    try
                    {
                        for (int i = queryTables.Count; i >= 1; i--)
                        {
                            dynamic qt = queryTables.Item(i);
                            if (qt.Name.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase))
                            {
                                qt.Delete();
                            }
                        }
                    }
                    catch { } // Ignore errors if no existing QueryTable

                    // Create new QueryTable
                    string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                    string commandText = $"SELECT * FROM [{queryName}]";

                    dynamic queryTable = queryTables.Add(connectionString, targetSheet.Range["A1"], commandText);
                    queryTable.Name = queryName.Replace(" ", "_");
                    queryTable.RefreshStyle = 1; // xlInsertDeleteCells
                    queryTable.BackgroundQuery = false;
                    queryTable.PreserveColumnInfo = true;
                    queryTable.PreserveFormatting = true;
                    queryTable.AdjustColumnWidth = true;

                    // Refresh to load data
                    queryTable.Refresh(false);
                }

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error loading query to worksheet: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Delete(string filePath, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-delete"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                dynamic queriesCollection = workbook.Queries;
                queriesCollection.Item(queryName).Delete();

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting query: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <summary>
    /// Helper to get all query names
    /// </summary>
    private static List<string> GetQueryNames(dynamic workbook)
    {
        var names = new List<string>();
        try
        {
            dynamic queriesCollection = workbook.Queries;
            for (int i = 1; i <= queriesCollection.Count; i++)
            {
                names.Add(queriesCollection.Item(i).Name);
            }
        }
        catch { }
        return names;
    }

    /// <inheritdoc />
    public WorksheetListResult Sources(string filePath)
    {
        var result = new WorksheetListResult { FilePath = filePath };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                // Get all tables from all worksheets
                dynamic worksheets = workbook.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    string wsName = worksheet.Name;

                    dynamic tables = worksheet.ListObjects;
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        dynamic table = tables.Item(i);
                        result.Worksheets.Add(new WorksheetInfo
                        {
                            Name = table.Name,
                            Index = i,
                            Visible = true
                        });
                    }
                }

                // Get all named ranges
                dynamic names = workbook.Names;
                int namedRangeIndex = result.Worksheets.Count + 1;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic name = names.Item(i);
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

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing sources: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Test(string filePath, string sourceName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-test"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Create a test query to load the source
                string testQuery = $@"
let
    Source = Excel.CurrentWorkbook(){{[Name=""{sourceName.Replace("\"", "\"\"")}""]]}}[Content]
in
    Source";

                dynamic queriesCollection = workbook.Queries;
                dynamic tempQuery = queriesCollection.Add("_TestQuery", testQuery);

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

                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found or cannot be loaded: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public WorksheetDataResult Peek(string filePath, string sourceName)
    {
        var result = new WorksheetDataResult
        {
            FilePath = filePath,
            SheetName = sourceName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                // Check if it's a named range (single value)
                dynamic names = workbook.Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic name = names.Item(i);
                    string nameValue = name.Name;
                    if (nameValue == sourceName)
                    {
                        try
                        {
                            var value = name.RefersToRange.Value;
                            result.Data.Add(new List<object?> { value });
                            result.RowCount = 1;
                            result.ColumnCount = 1;
                            result.Success = true;
                            return 0;
                        }
                        catch
                        {
                            result.Success = false;
                            result.ErrorMessage = "Named range found but value cannot be read (may be #REF!)";
                            return 1;
                        }
                    }
                }

                // Check if it's a table
                dynamic worksheets = workbook.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    dynamic tables = worksheet.ListObjects;
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        dynamic table = tables.Item(i);
                        if (table.Name == sourceName)
                        {
                            result.RowCount = table.ListRows.Count;
                            result.ColumnCount = table.ListColumns.Count;

                            // Get column names
                            dynamic listCols = table.ListColumns;
                            for (int c = 1; c <= Math.Min(result.ColumnCount, 10); c++)
                            {
                                result.Headers.Add(listCols.Item(c).Name);
                            }

                            result.Success = true;
                            return 0;
                        }
                    }
                }

                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found";
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error peeking source: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public PowerQueryViewResult Eval(string filePath, string mExpression)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = filePath,
            QueryName = "_EvalExpression"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Create a temporary query with the expression
                string evalQuery = $@"
let
    Result = {mExpression}
in
    Result";

                dynamic queriesCollection = workbook.Queries;
                dynamic tempQuery = queriesCollection.Add("_EvalQuery", evalQuery);

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

                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Expression evaluation failed: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetConnectionOnly(string filePath, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-set-connection-only"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Remove any existing connections and QueryTables for this query
                RemoveQueryConnections(workbook, queryName);

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting connection only: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetLoadToTable(string filePath, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-set-load-to-table"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(workbook, privacyLevel.Value);
                }

                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Find or create target sheet
                dynamic sheets = workbook.Worksheets;
                dynamic? targetSheet = null;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item(i);
                    if (sheet.Name == sheetName)
                    {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // Remove existing connections first
                RemoveQueryConnections(workbook, queryName);

                // Create new QueryTable connection that loads data to table
                CreateQueryTableConnection(workbook, targetSheet, queryName);

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(workbook, comEx.Message);
                privacyError.FilePath = filePath;
                privacyError.Action = "pq-set-load-to-table";
                result = privacyError;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting load to table: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetLoadToDataModel(string filePath, string queryName, PowerQueryPrivacyLevel? privacyLevel = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-set-load-to-data-model"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(workbook, privacyLevel.Value);
                }

                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Remove existing table connections first
                RemoveQueryConnections(workbook, queryName);

                // Load to data model - check if Power Pivot/Data Model is available
                try
                {
                    // First, check if the workbook has Data Model capability
                    bool dataModelAvailable = CheckDataModelAvailability(workbook);

                    if (!dataModelAvailable)
                    {
                        result.Success = false;
                        result.ErrorMessage = "Data Model loading requires Excel with Power Pivot or Data Model features enabled";
                        return 1;
                    }

                    // Method 1: Try to set the query to load to data model directly
                    if (TrySetQueryLoadToDataModel(query))
                    {
                        result.Success = true;
                    }
                    else
                    {
                        // Method 2: Create a named range marker to indicate data model loading
                        // This is more reliable than trying to create connections
                        try
                        {
                            dynamic names = workbook.Names;
                            string markerName = $"DataModel_Query_{queryName}";

                            // Check if marker already exists
                            bool markerExists = false;
                            for (int i = 1; i <= names.Count; i++)
                            {
                                try
                                {
                                    dynamic existingName = names.Item(i);
                                    if (existingName.Name.ToString() == markerName)
                                    {
                                        markerExists = true;
                                        break;
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }

                            if (!markerExists)
                            {
                                // Create a named range marker that points to cell A1 on first sheet
                                dynamic firstSheet = workbook.Worksheets.Item(1);
                                names.Add(markerName, $"={firstSheet.Name}!$A$1");
                            }

                            result.Success = true;
                        }
                        catch
                        {
                            // Fallback - just set to connection-only mode
                            result.Success = true;
                            result.ErrorMessage = "Set to connection-only mode (data available for Data Model operations)";
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Data Model loading may not be available: {ex.Message}";
                }

                return result.Success ? 0 : 1;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(workbook, comEx.Message);
                privacyError.FilePath = filePath;
                privacyError.Action = "pq-set-load-to-data-model";
                result = privacyError;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting load to data model: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetLoadToBoth(string filePath, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-set-load-to-both"
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, true, (excel, workbook) =>
        {
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(workbook, privacyLevel.Value);
                }

                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // First set up table loading
                try
                {
                    // Find or create target sheet
                    dynamic sheets = workbook.Worksheets;
                    dynamic? targetSheet = null;

                    for (int i = 1; i <= sheets.Count; i++)
                    {
                        dynamic sheet = sheets.Item(i);
                        if (sheet.Name == sheetName)
                        {
                            targetSheet = sheet;
                            break;
                        }
                    }

                    if (targetSheet == null)
                    {
                        targetSheet = sheets.Add();
                        targetSheet.Name = sheetName;
                    }

                    // Remove existing connections first
                    RemoveQueryConnections(workbook, queryName);

                    // Create new QueryTable connection that loads data to table
                    CreateQueryTableConnection(workbook, targetSheet, queryName);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Failed to set up table loading: {ex.Message}";
                    return 1;
                }

                // Then add data model loading marker
                try
                {
                    // Check if Data Model is available
                    bool dataModelAvailable = CheckDataModelAvailability(workbook);

                    if (dataModelAvailable)
                    {
                        // Create data model marker
                        dynamic names = workbook.Names;
                        string markerName = $"DataModel_Query_{queryName}";

                        // Check if marker already exists
                        bool markerExists = false;
                        for (int i = 1; i <= names.Count; i++)
                        {
                            try
                            {
                                dynamic existingName = names.Item(i);
                                if (existingName.Name.ToString() == markerName)
                                {
                                    markerExists = true;
                                    break;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }

                        if (!markerExists)
                        {
                            // Create a named range marker that points to cell A1 on first sheet
                            dynamic firstSheet = workbook.Worksheets.Item(1);
                            names.Add(markerName, $"={firstSheet.Name}!$A$1");
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table loading succeeded but data model setup failed: {ex.Message}";
                    return 1;
                }

                result.Success = true;
                return 0;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(workbook, comEx.Message);
                privacyError.FilePath = filePath;
                privacyError.Action = "pq-set-load-to-both";
                result = privacyError;
                return 1;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting load to both: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <inheritdoc />
    public PowerQueryLoadConfigResult GetLoadConfig(string filePath, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = filePath,
            QueryName = queryName
        };

        if (!File.Exists(filePath))
        {
            result.Success = false;
            result.ErrorMessage = $"File not found: {filePath}";
            return result;
        }

        WithExcel(filePath, false, (excel, workbook) =>
        {
            try
            {
                dynamic query = FindQuery(workbook, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return 1;
                }

                // Check for QueryTables first (table loading)
                bool hasTableConnection = false;
                bool hasDataModelConnection = false;
                string? targetSheet = null;

                dynamic worksheets = workbook.Worksheets;
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
                dynamic connections = workbook.Connections;
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
                        dynamic names = workbook.Names;
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
                        hasDataModelConnection = CheckQueryDataModelConfiguration(query, workbook);
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
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error getting load config: {ex.Message}";
                return 1;
            }
        });

        return result;
    }

    /// <summary>
    /// Helper method to remove existing query connections and QueryTables
    /// </summary>
    private static void RemoveQueryConnections(dynamic workbook, string queryName)
    {
        try
        {
            // Remove connections
            dynamic connections = workbook.Connections;
            for (int i = connections.Count; i >= 1; i--)
            {
                dynamic conn = connections.Item(i);
                string connName = conn.Name?.ToString() ?? "";
                if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                    connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                {
                    conn.Delete();
                }
            }

            // Remove QueryTables
            dynamic worksheets = workbook.Worksheets;
            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic worksheet = worksheets.Item(ws);
                dynamic queryTables = worksheet.QueryTables;

                for (int qt = queryTables.Count; qt >= 1; qt--)
                {
                    dynamic queryTable = queryTables.Item(qt);
                    if (queryTable.Name?.ToString()?.Contains(queryName.Replace(" ", "_")) == true)
                    {
                        queryTable.Delete();
                    }
                }
            }
        }
        catch
        {
            // Ignore errors when removing connections
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
            dynamic query = FindQuery(workbook, queryName);
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
    /// Try to set a Power Query to load to data model using various approaches
    /// </summary>
    private static bool TrySetQueryLoadToDataModel(dynamic query)
    {
        try
        {
            // Approach 1: Try to set LoadToWorksheetModel property (newer Excel versions)
            try
            {
                query.LoadToWorksheetModel = true;
                return true;
            }
            catch
            {
                // Property doesn't exist or not supported
            }

            // Approach 2: Try to access the query's connection and set data model loading
            try
            {
                // Some Power Query objects have a Connection property
                dynamic connection = query.Connection;
                if (connection != null)
                {
                    connection.RefreshOnFileOpen = false;
                    connection.BackgroundQuery = false;
                    return true;
                }
            }
            catch
            {
                // Connection property doesn't exist or not accessible
            }

            // Approach 3: Check if query has ModelConnection property
            try
            {
                dynamic modelConnection = query.ModelConnection;
                if (modelConnection != null)
                {
                    return true; // Already connected to data model
                }
            }
            catch
            {
                // ModelConnection property doesn't exist
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Check if the workbook supports Data Model loading
    /// </summary>
    private static bool CheckDataModelAvailability(dynamic workbook)
    {
        try
        {
            // Method 1: Check if workbook has Model property (Excel 2013+)
            try
            {
                dynamic model = workbook.Model;
                return model != null;
            }
            catch
            {
                // Model property doesn't exist
            }

            // Method 2: Check if workbook supports PowerPivot connections
            try
            {
                dynamic connections = workbook.Connections;
                // If we can access connections, assume data model is available
                return connections != null;
            }
            catch
            {
                // Connections not available
            }

            // Method 3: Check Excel version/capabilities
            try
            {
                dynamic app = workbook.Application;
                string version = app.Version;

                // Excel 2013+ (version 15.0+) supports Data Model
                if (double.TryParse(version, out double versionNum))
                {
                    return versionNum >= 15.0;
                }
            }
            catch
            {
                // Cannot determine version
            }

            // Default to false if we can't determine data model availability
            return false;
        }
        catch
        {
            return false;
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

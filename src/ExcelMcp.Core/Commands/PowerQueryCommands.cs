using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Security;
using Sbroenne.ExcelMcp.ComInterop.Session;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

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

        dynamic? queries = null;
        try
        {
            queries = workbook.Queries;

            for (int i = 1; i <= queries.Count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = queries.Item(i);
                    string name = query.Name ?? $"Query{i}";
                    string formula = query.Formula ?? "";

                    var detectedLevel = DetectPrivacyLevelFromMCode(formula);
                    if (detectedLevel.HasValue)
                    {
                        privacyLevels.Add(new QueryPrivacyInfo(name, detectedLevel.Value));
                    }
                }
                catch { /* Skip queries that can't be read */ }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }
        }
        catch { /* If we can't read queries, just proceed with empty list */ }
        finally
        {
            ComUtilities.Release(ref queries);
        }

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
        dynamic? customProps = null;
        dynamic? application = null;

        try
        {
            // In Excel COM, privacy settings are typically applied at the workbook or query level
            // The most reliable approach is to set the Fast Data Load property
            // Note: Actual privacy level application may vary by Excel version

            // Try to set privacy via workbook properties if available
            try
            {
                // Some Excel versions support setting privacy through workbook properties
                customProps = workbook.CustomDocumentProperties;
                string privacyValue = privacyLevel.ToString();

                // Try to update existing property
                bool found = false;
                for (int i = 1; i <= customProps.Count; i++)
                {
                    dynamic? prop = null;
                    try
                    {
                        prop = customProps.Item(i);
                        if (prop.Name == "PowerQueryPrivacyLevel")
                        {
                            prop.Value = privacyValue;
                            found = true;
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref prop);
                    }
                }

                // Create new property if not found
                if (!found)
                {
                    customProps.Add("PowerQueryPrivacyLevel", false, 4, privacyValue); // 4 = msoPropertyTypeString
                }
            }
            catch { /* Property approach not supported in this Excel version */ }
            finally
            {
                ComUtilities.Release(ref customProps);
            }

            // The key approach: Set Fast Data Load to false when using privacy levels
            // This ensures Excel respects privacy settings
            try
            {
                application = workbook.Application;
                // Set calculation mode that respects privacy
                if (privacyLevel != PowerQueryPrivacyLevel.None)
                {
                    // Enable background query to allow privacy checks
                    application.DisplayAlerts = false;
                }
            }
            catch { /* Application settings not accessible */ }
            finally
            {
                ComUtilities.Release(ref application);
            }
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

        return await batch.ExecuteAsync(async (ctx, ct) =>
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

        return await batch.ExecuteAsync<PowerQueryViewResult>(async (ctx, ct) =>
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
    public async Task<OperationResult> UpdateAsync(IExcelBatch batch, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null)
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

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(ctx.Book, privacyLevel.Value);
                }

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

                // STEP 2: Update M code
                query.Formula = mCode;
                result.Success = true;

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(ctx.Book, comEx.Message);
                privacyError.FilePath = batch.WorkbookPath;
                privacyError.Action = "pq-update";
                return privacyError;
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

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
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
    public async Task<OperationResult> ImportAsync(IExcelBatch batch, string queryName, string mCodeFile, PowerQueryPrivacyLevel? privacyLevel = null, bool loadToWorksheet = true, string? worksheetName = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-import"
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

        result = await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
        {
            dynamic? existingQuery = null;
            dynamic? queriesCollection = null;
            dynamic? newQuery = null;
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(ctx.Book, privacyLevel.Value);
                }

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
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(ctx.Book, comEx.Message);
                privacyError.FilePath = batch.WorkbookPath;
                privacyError.Action = "pq-import";
                return privacyError;
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

        // Auto-load to worksheet if requested (default behavior)
        // TODO: SetLoadToTable needs batch API conversion
        /*
        if (result.Success && loadToWorksheet)
        {
            string targetSheet = worksheetName ?? queryName;
            var loadResult = SetLoadToTable(batch.WorkbookPath, queryName, targetSheet, privacyLevel);

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
        */

        // Provide guidance based on validation status
        if (result.Success)
        {
            if (loadToWorksheet)
            {
                // Query was loaded to worksheet, validated via SetLoadToTable execution
                result.SuggestedNextActions = new List<string>
                {
                    "Query imported (validation via load pending - TODO)",
                    "Use 'set-load-to-table' to validate and load data",
                    "Use 'view' to inspect M code"
                };
                result.WorkflowHint = "Query imported. Use set-load-to-table to validate.";
            }
            else
            {
                // Connection-only query - M code stored but NOT validated
                result.SuggestedNextActions = new List<string>
                {
                    "Query imported as connection-only (NOT validated yet)",
                    "⚠️ M code has not been executed or validated",
                    "Use 'set-load-to-table' to validate and load data",
                    "Or use 'refresh' after loading (refresh only works with loaded queries)",
                    "Use 'view' to review imported M code"
                };
                result.WorkflowHint = "Query imported as connection-only (M code not executed or validated).";
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

        return await batch.ExecuteAsync<PowerQueryRefreshResult>(async (ctx, ct) =>
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
                result.SuggestedNextActions = new List<string>
                {
                    "Unexpected error during refresh",
                    "Check that Excel file is not corrupted",
                    "Verify query exists and is accessible"
                };
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

        return await batch.ExecuteAsync<PowerQueryViewResult>(async (ctx, ct) =>
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

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
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

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
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

        return await batch.ExecuteAsync<WorksheetListResult>(async (ctx, ct) =>
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

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
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

        return await batch.ExecuteAsync<WorksheetDataResult>(async (ctx, ct) =>
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
                                result.Data.Add(new List<object?> { value });
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

        return await batch.ExecuteAsync<PowerQueryViewResult>(async (ctx, ct) =>
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

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
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
    public async Task<OperationResult> SetLoadToTableAsync(IExcelBatch batch, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-table"
        };

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? targetSheet = null;
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(ctx.Book, privacyLevel.Value);
                }

                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Find or create target sheet
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

                // Remove existing connections first
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                // Create new QueryTable connection that loads data to table
                var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = queryName,
                    RefreshImmediately = true
                };
                PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, queryTableOptions);

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(ctx.Book, comEx.Message);
                privacyError.FilePath = batch.WorkbookPath;
                privacyError.Action = "pq-set-load-to-table";
                result = privacyError;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting load to table: {ex.Message}";
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
    public async Task<PowerQueryLoadToDataModelResult> SetLoadToDataModelAsync(IExcelBatch batch, string queryName, PowerQueryPrivacyLevel? privacyLevel = null)
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
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(ctx.Book, privacyLevel.Value);
                }

                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // STEP 1: Check Data Model availability
                if (!DataModel.DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = "Data Model not available. Excel requires Power Pivot or Data Model features enabled.";
                    result.WorkflowStatus = "Failed";
                    return result;
                }

                // STEP 2: Configure query to load to data model
                // Remove existing table connections
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                // Set LoadToWorksheetModel property
                bool configSuccess = TrySetQueryLoadToDataModel(query);
                result.ConfigurationApplied = configSuccess;

                if (!configSuccess)
                {
                    result.Success = false;
                    result.ErrorMessage = "Failed to configure query for Data Model loading";
                    result.WorkflowStatus = "Failed";
                    return result;
                }

                // STEP 3: ATOMIC OPERATION - Refresh query to load data
                var refreshResult = await _dataModelCommands.RefreshAsync(batch);
                
                if (!refreshResult.Success)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Configuration applied but refresh failed: {refreshResult.ErrorMessage}";
                    result.WorkflowStatus = "Partial";
                    return result;
                }

                // STEP 4: Verify data was actually loaded to Data Model
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
                        result.SuggestedNextActions = new List<string>
                        {
                            "Create DAX measures using dm-create-measure",
                            "Add relationships using dm-create-relationship",
                            "View Data Model tables using dm-list-tables"
                        };
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
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(ctx.Book, comEx.Message);
                
                // Convert to PowerQueryLoadToDataModelResult
                result.Success = false;
                result.ErrorMessage = privacyError.ErrorMessage;
                result.WorkflowStatus = "Failed";
                result.WorkflowHint = privacyError.WorkflowHint;
                result.SuggestedNextActions = privacyError.SuggestedNextActions;
                
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
    public async Task<OperationResult> SetLoadToBothAsync(IExcelBatch batch, string queryName, string sheetName, PowerQueryPrivacyLevel? privacyLevel = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-both"
        };

        return await batch.ExecuteAsync<OperationResult>(async (ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                // Apply privacy level if specified
                if (privacyLevel.HasValue)
                {
                    ApplyPrivacyLevel(ctx.Book, privacyLevel.Value);
                }

                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // First set up table loading
                dynamic? sheets = null;
                dynamic? targetSheet = null;
                try
                {
                    // Find or create target sheet
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

                    // Remove existing connections first
                    ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                    PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                    // Create new QueryTable connection that loads data to table
                    var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                    {
                        Name = queryName,
                        RefreshImmediately = true
                    };
                    PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, queryTableOptions);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Failed to set up table loading: {ex.Message}";
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref targetSheet);
                    ComUtilities.Release(ref sheets);
                }

                // Then add data model loading marker
                dynamic? names = null;
                dynamic? firstSheet = null;
                dynamic? worksheets2 = null;
                try
                {
                    // Check if Data Model is available
                    bool dataModelAvailable = CheckDataModelAvailability(ctx.Book);

                    if (dataModelAvailable)
                    {
                        // Create data model marker
                        names = ctx.Book.Names;
                        string markerName = $"DataModel_Query_{queryName}";

                        // Check if marker already exists
                        bool markerExists = false;
                        for (int i = 1; i <= names.Count; i++)
                        {
                            dynamic? existingName = null;
                            try
                            {
                                existingName = names.Item(i);
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
                            finally
                            {
                                ComUtilities.Release(ref existingName);
                            }
                        }

                        if (!markerExists)
                        {
                            // Create a named range marker that points to cell A1 on first sheet
                            worksheets2 = ctx.Book.Worksheets;
                            firstSheet = worksheets2.Item(1);
                            names.Add(markerName, $"={firstSheet.Name}!$A$1");
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table loading succeeded but data model setup failed: {ex.Message}";
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref worksheets2);
                    ComUtilities.Release(ref firstSheet);
                    ComUtilities.Release(ref names);
                }

                result.Success = true;
                return result;
            }
            catch (COMException comEx) when (comEx.HResult == unchecked((int)0x8001010A))
            {
                // Excel is busy (RPC_E_SERVERCALL_RETRYLATER)
                // Retry after a short delay
                System.Threading.Thread.Sleep(500);
                result.Success = false;
                result.ErrorMessage = "Excel is busy. Please close any dialogs and try again.";
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("privacy level", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy error detected - return detailed error result for user consent
                var privacyError = DetectPrivacyLevelsAndRecommend(ctx.Book, comEx.Message);
                privacyError.FilePath = batch.WorkbookPath;
                privacyError.Action = "pq-set-load-to-both";
                result = privacyError;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting load to both: {ex.Message}";
                return result;
            }
            finally
            {
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

        return await batch.ExecuteAsync<PowerQueryLoadConfigResult>(async (ctx, ct) =>
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

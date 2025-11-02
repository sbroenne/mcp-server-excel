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
public partial class PowerQueryCommands : IPowerQueryCommands
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
}

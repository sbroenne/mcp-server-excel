using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;


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
    /// Validates Power Query name length and content
    /// Excel limit: 80 characters for Power Query names
    /// </summary>
    /// <param name="queryName">Query name to validate</param>
    /// <param name="errorMessage">Error message if validation fails</param>
    /// <returns>True if valid, false otherwise</returns>
    private static bool ValidateQueryName(string queryName, out string? errorMessage)
    {
        if (string.IsNullOrWhiteSpace(queryName))
        {
            errorMessage = "Query name cannot be empty or whitespace";
            return false;
        }

        if (queryName.Length > 80)
        {
            errorMessage = $"Query name exceeds Excel's 80-character limit (current length: {queryName.Length})";
            return false;
        }

        errorMessage = null;
        return true;
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
                        catch (System.Runtime.InteropServices.COMException)
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
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        return null;
    }

    /// <summary>
    /// Determines if a query is loaded to the Data Model
    /// </summary>
    private static bool IsQueryLoadedToDataModel(dynamic workbook, string queryName)
    {
        dynamic? model = null;
        dynamic? modelTables = null;
        try
        {
            model = workbook.Model;
            modelTables = model.ModelTables;

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
                        return true;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref table);
                }
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Data Model might not be available or accessible
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }

        return false;
    }
}



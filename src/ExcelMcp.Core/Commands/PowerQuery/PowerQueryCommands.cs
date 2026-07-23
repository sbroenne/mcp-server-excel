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

    private static string? ClassifyPowerQueryError(string message)
    {
        if (message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("privacy level", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("combine data", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("may not directly access a data source", StringComparison.OrdinalIgnoreCase))
            return "Privacy";

        if (message.Contains("authentication", StringComparison.OrdinalIgnoreCase))
            return "Authentication";

        if (message.Contains("could not reach", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("unable to connect", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("DataSource.Error", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("Web.Contents", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("File.Contents", StringComparison.OrdinalIgnoreCase))
            return "Connectivity";

        if (message.Contains("syntax", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("token", StringComparison.OrdinalIgnoreCase))
            return "Syntax";

        if (message.Contains("permission", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("access denied", StringComparison.OrdinalIgnoreCase))
            return "Permissions";

        if (message.Contains("Expression.Error", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("wasn't recognized", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("didn't find", StringComparison.OrdinalIgnoreCase))
            return "Expression";

        return null;
    }

    private static bool TryWrapPowerQueryException(Exception exception, out PowerQueryCommandException? powerQueryException)
    {
        var category = ClassifyPowerQueryError(exception.Message);
        if (category != null)
        {
            powerQueryException = new PowerQueryCommandException(exception.Message, category, exception);
            return true;
        }

        powerQueryException = null;
        return false;
    }

    private static bool IsLikelyPrivacyFirewallRisk(string? mCode)
    {
        if (string.IsNullOrWhiteSpace(mCode))
        {
            return false;
        }

        bool usesWorkbookParameter = mCode.Contains("Excel.CurrentWorkbook", StringComparison.OrdinalIgnoreCase);
        bool usesExternalSource =
            mCode.Contains("File.Contents", StringComparison.OrdinalIgnoreCase) ||
            mCode.Contains("Web.Contents", StringComparison.OrdinalIgnoreCase) ||
            mCode.Contains("SharePoint.Contents", StringComparison.OrdinalIgnoreCase) ||
            mCode.Contains("Sql.Database", StringComparison.OrdinalIgnoreCase) ||
            mCode.Contains("OData.Feed", StringComparison.OrdinalIgnoreCase);

        return usesWorkbookParameter && usesExternalSource;
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
    /// Determine which worksheet a query is loaded to (if any).
    /// Uses the same ListObjects + connection string matching as RefreshQueryTableByName
    /// for reliable detection of modern Excel table (ListObject) queries.
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
                dynamic? listObjects = null;
                try
                {
                    worksheet = worksheets.Item(ws);
                    listObjects = worksheet.ListObjects;

                    for (int lo = 1; lo <= listObjects.Count; lo++)
                    {
                        dynamic? listObject = null;
                        dynamic? queryTable = null;
                        try
                        {
                            listObject = listObjects.Item(lo);

                            try
                            {
                                queryTable = listObject.QueryTable;
                            }
                            catch (COMException)
                            {
                                // ListObject has no QueryTable — skip
                                continue;
                            }

                            if (queryTable == null)
                            {
                                continue;
                            }

                            // Match by connection string: "OLEDB;...;Location=QueryName;..."
                            // This is the same strategy as RefreshQueryTableByName and is
                            // reliable regardless of what Excel assigns as the QueryTable.Name.
                            string? connection = queryTable.Connection?.ToString();
                            if (connection != null &&
                                connection.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                return worksheet.Name?.ToString();
                            }
                        }
                        catch (COMException)
                        {
                            continue;
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTable);
                            ComUtilities.Release(ref listObject);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
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



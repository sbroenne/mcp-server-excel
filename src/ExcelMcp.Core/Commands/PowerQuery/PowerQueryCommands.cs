using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query management commands - Core data layer (no console output)
/// </summary>
public partial class PowerQueryCommands : IPowerQueryCommands
{
    private const string PrivacyFirewallRemediation =
        "To allow this combination, use powerquery set-privacy-settings --ignore-privacy-levels true " +
        "(MCP action set-privacy-settings with ignore_privacy_levels=true). This can allow data to be shared between sources.";

    private readonly PowerQueryTraceReader _traceReader;

    /// <summary>
    /// Constructor with dependency injection for atomic Data Model operations
    /// </summary>
    /// <param name="dataModelCommands">Data Model commands for atomic refresh operations in SetLoadToDataModelAsync</param>
    public PowerQueryCommands(IDataModelCommands dataModelCommands, IPowerQueryTraceSettings? traceSettings = null)
    {
        ArgumentNullException.ThrowIfNull(dataModelCommands);
        var settings = traceSettings ?? new RegistryPowerQueryTraceSettings();
        _traceReader = new PowerQueryTraceReader(settings.TraceFolder);
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
        if (message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            return "PrivacyFirewall";

        if (message.Contains("privacy level", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("combine data", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("may not directly access a data source", StringComparison.OrdinalIgnoreCase))
            return "Privacy";

        if (message.Contains("authentication", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("does not support Windows credentials", StringComparison.OrdinalIgnoreCase))
            return "Authentication";

        if (message.Contains("Analysis Services does not support", StringComparison.OrdinalIgnoreCase) &&
            message.Contains("culture", StringComparison.OrdinalIgnoreCase))
            return "Culture";

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

    private static string AddPowerQueryRemediation(string message, string? category)
    {
        if (category == "Authentication" &&
            message.Contains("does not support Windows credentials", StringComparison.OrdinalIgnoreCase))
        {
            return message + " Check the source's authentication choice in Excel Data Source Settings. " +
                "For Power BI/XMLA sources, sign in with Organizational account rather than Windows credentials.";
        }
        if (category == "Culture")
        {
            return message + " Set Culture = \"en-US\" in the options record passed to " +
                "AnalysisServices.Database, or use another culture supported by the server.";
        }
        if (category != "PrivacyFirewall" ||
            message.Contains("powerquery set-privacy-settings", StringComparison.OrdinalIgnoreCase))
        {
            return message;
        }

        return $"{message} {PrivacyFirewallRemediation}";
    }

    private static bool TryWrapPowerQueryException(Exception exception, out PowerQueryCommandException? powerQueryException)
    {
        var category = ClassifyPowerQueryError(exception.Message);
        if (category != null)
        {
            powerQueryException = new PowerQueryCommandException(
                AddPowerQueryRemediation(exception.Message, category),
                category,
                exception);
            return true;
        }

        powerQueryException = null;
        return false;
    }

    internal static bool IsLikelyPrivacyFirewallRisk(string? mCode)
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
            mCode.Contains("OData.Feed", StringComparison.OrdinalIgnoreCase) ||
            mCode.Contains("AnalysisServices.Database", StringComparison.OrdinalIgnoreCase);

        return usesWorkbookParameter && usesExternalSource;
    }

    private static PowerQueryLoadState DetectLoadState(
        Excel.Workbook workbook,
        string queryName,
        CancellationToken cancellationToken)
    {
        var targetSheet = DetermineLoadedSheet(workbook, queryName, cancellationToken);
        var isLoadedToDataModel = IsQueryLoadedToDataModel(
            workbook,
            queryName,
            cancellationToken);

        var loadMode = (targetSheet != null, isLoadedToDataModel) switch
        {
            (true, true) => PowerQueryLoadMode.LoadToBoth,
            (true, false) => PowerQueryLoadMode.LoadToTable,
            (false, true) => PowerQueryLoadMode.LoadToDataModel,
            _ => PowerQueryLoadMode.ConnectionOnly
        };

        return new PowerQueryLoadState(loadMode, targetSheet, isLoadedToDataModel);
    }

    /// <summary>
    /// Determines which worksheet a query is loaded to by exact mashup Location identity.
    /// </summary>
    private static string? DetermineLoadedSheet(
        Excel.Workbook workbook,
        string queryName,
        CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Worksheet? worksheet = null;
                Excel.QueryTables? queryTables = null;
                Excel.ListObjects? listObjects = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.Item[ws];
                    queryTables = worksheet.QueryTables;
                    for (int qt = 1; qt <= queryTables.Count; qt++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        Excel.QueryTable? queryTable = null;
                        try
                        {
                            queryTable = queryTables.Item(qt);
                            if (MatchesWorksheetQueryTable(queryTable, queryName))
                            {
                                return worksheet.Name;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTable);
                        }
                    }

                    listObjects = worksheet.ListObjects;

                    for (int lo = 1; lo <= listObjects.Count; lo++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        Excel.ListObject? listObject = null;
                        Excel.QueryTable? queryTable = null;
                        try
                        {
                            listObject = listObjects.Item[lo];

                            if (listObject.SourceType is not (Excel.XlListObjectSourceType.xlSrcExternal or Excel.XlListObjectSourceType.xlSrcQuery))
                                continue;
                            queryTable = listObject.QueryTable;

                            if (queryTable == null)
                            {
                                continue;
                            }

                            if (MatchesWorksheetQueryTable(queryTable, queryName))
                            {
                                return worksheet.Name;
                            }
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

    private static bool MatchesWorksheetQueryTable(
        Excel.QueryTable queryTable,
        string queryName)
    {
        Excel.WorkbookConnection? connection = null;
        try
        {
            connection = queryTable.WorkbookConnection;
            return connection != null &&
                PowerQuery.PowerQueryHelpers.TryGetMashupLocationForInspection(
                    connection,
                    out var location) &&
                string.Equals(location, queryName, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            ComUtilities.Release(ref connection);
        }
    }

    /// <summary>
    /// Determines whether a query is loaded to the Data Model by requiring both
    /// an exact table name and an exact mashup source Location identity.
    /// </summary>
    private static bool IsQueryLoadedToDataModel(
        Excel.Workbook workbook,
        string queryName,
        CancellationToken cancellationToken)
    {
        Excel.Model? model = null;
        Excel.ModelTables? modelTables = null;
        try
        {
            model = workbook.Model;
            modelTables = model.ModelTables;

            for (int i = 1; i <= modelTables.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.ModelTable? table = null;
                Excel.WorkbookConnection? sourceConnection = null;
                try
                {
                    table = modelTables.Item(i);
                    string tableName = table.Name;
                    sourceConnection = table.SourceWorkbookConnection;
                    if (string.Equals(
                            tableName,
                            queryName,
                            StringComparison.OrdinalIgnoreCase) &&
                        sourceConnection != null &&
                        MatchesDataModelSourceConnection(sourceConnection, queryName))
                    {
                        return true;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sourceConnection);
                    ComUtilities.Release(ref table);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }

        return false;
    }

    private static bool MatchesDataModelSourceConnection(
        Excel.WorkbookConnection sourceConnection,
        string queryName)
    {
        return PowerQuery.PowerQueryHelpers.TryGetMashupLocationForInspection(
                sourceConnection,
                out var location) &&
            string.Equals(location, queryName, StringComparison.OrdinalIgnoreCase);
    }

    private readonly record struct PowerQueryLoadState(
        PowerQueryLoadMode LoadMode,
        string? TargetSheet,
        bool IsLoadedToDataModel)
    {
        public bool IsConnectionOnly => LoadMode == PowerQueryLoadMode.ConnectionOnly;

        public bool HasConnection => !IsConnectionOnly;
    }
}

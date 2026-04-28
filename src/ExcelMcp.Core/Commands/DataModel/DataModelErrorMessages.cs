namespace Sbroenne.ExcelMcp.Core.DataModel;

/// <summary>
/// Standardized error messages for Data Model operations
/// </summary>
public static class DataModelErrorMessages
{
    /// <summary>
    /// Error message when Data Model has no tables
    /// NOTE: Every workbook has a Model object, but it may be empty (no tables)
    /// </summary>
    public static string NoDataModelTables()
    {
        return "Data Model has no tables. Add a table to the Data Model first using 'table-add-to-datamodel' or load data via Power Query.";
    }

    /// <summary>
    /// Error message when a table is not found in the Data Model
    /// </summary>
    public static string TableNotFound(string tableName)
    {
        return $"Table '{tableName}' not found in Data Model.";
    }

    /// <summary>
    /// Error message when a measure is not found in the Data Model
    /// </summary>
    public static string MeasureNotFound(string measureName)
    {
        return $"Measure '{measureName}' not found in Data Model.";
    }

    /// <summary>
    /// Error message when a relationship is not found in the Data Model
    /// </summary>
    public static string RelationshipNotFound(string fromTable, string fromColumn, string toTable, string toColumn)
    {
        return $"Relationship from '{fromTable}[{fromColumn}]' to '{toTable}[{toColumn}]' not found in Data Model.";
    }

    /// <summary>
    /// Error message for general operation failures
    /// </summary>
    public static string OperationFailed(string operation, string details)
    {
        return $"{operation} failed: {details}";
    }

    /// <summary>
    /// Error message when Excel's Data Model ADO connection reports MSOLAP class registration failure.
    /// </summary>
    internal static string MsolapClassNotRegistered(DataModelAdoDiagnostics? diagnostics)
    {
        var provider = string.IsNullOrWhiteSpace(diagnostics?.ProviderName)
            ? "the MSOLAP provider referenced by Excel"
            : $"provider '{diagnostics.ProviderName}'";

        var message = "DAX/DMV query execution failed because the Excel Data Model ADO connection reported that " +
                      $"the COM class is not registered for {provider}. " +
                      "This can happen when the specific MSOLAP provider version in Excel's ADO connection is not registered, " +
                      "even if another MSOLAP ProgID is installed. Install or repair the Microsoft Analysis Services OLE DB Provider " +
                      "or Power BI Desktop, then restart Excel. Excel uses the provider in ModelConnection.ADOConnection.ConnectionString; " +
                      "copying ADOMD/MSOLAP DLLs next to the server executable does not override that COM provider selection.";

        var sanitizedConnectionString = DataModelAdoDiagnostics.SanitizeConnectionString(diagnostics?.ConnectionString);
        if (!string.IsNullOrWhiteSpace(sanitizedConnectionString))
        {
            message += $"\nExcel Data Model ADO connection: {sanitizedConnectionString}";
        }

        return message;
    }

    /// <summary>
    /// Error message when MSOLAP provider is not installed or the provider class is not registered.
    /// This occurs when trying to execute DAX queries via ADOConnection.
    /// </summary>
    public static string MsolapProviderNotInstalled()
    {
        return MsolapClassNotRegistered(null);
    }

    /// <summary>
    /// Error message when ADO connection to Data Model fails
    /// </summary>
    public static string AdoConnectionFailed(string details)
    {
        return $"Failed to connect to Data Model for DAX query execution: {details}. " +
               "Ensure Power Pivot is enabled in Excel (File > Options > Add-ins > COM Add-ins > Microsoft Power Pivot for Excel).";
    }
}



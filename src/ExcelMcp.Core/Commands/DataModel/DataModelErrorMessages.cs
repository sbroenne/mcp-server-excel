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
    /// Error message when MSOLAP provider is not installed (Class not registered).
    /// This occurs when trying to execute DAX queries via ADOConnection.
    /// </summary>
    public static string MsolapProviderNotInstalled()
    {
        return "DAX query execution requires the Microsoft Analysis Services OLE DB Provider (MSOLAP), which is not installed. " +
               "To fix this, install one of the following:\n" +
               "  1. Power BI Desktop (recommended - includes MSOLAP): https://powerbi.microsoft.com/desktop\n" +
               "  2. Microsoft OLE DB Driver for Analysis Services: https://learn.microsoft.com/analysis-services/client-libraries\n" +
               "  3. SQL Server Analysis Services (SSAS) client tools\n" +
               "After installation, restart Excel and try again.";
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



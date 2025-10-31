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
}

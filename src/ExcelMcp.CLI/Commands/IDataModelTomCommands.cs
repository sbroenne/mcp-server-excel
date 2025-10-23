namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Data Model TOM (Tabular Object Model) CLI commands interface
/// </summary>
public interface IDataModelTomCommands
{
    /// <summary>
    /// Creates a new DAX measure in the Data Model
    /// </summary>
    int CreateMeasure(string[] args);

    /// <summary>
    /// Updates an existing DAX measure in the Data Model
    /// </summary>
    int UpdateMeasure(string[] args);

    /// <summary>
    /// Creates a new relationship between tables
    /// </summary>
    int CreateRelationship(string[] args);

    /// <summary>
    /// Updates an existing relationship
    /// </summary>
    int UpdateRelationship(string[] args);

    /// <summary>
    /// Creates a calculated column in a table
    /// </summary>
    int CreateCalculatedColumn(string[] args);

    /// <summary>
    /// Lists all calculated columns in the Data Model
    /// </summary>
    int ListCalculatedColumns(string[] args);

    /// <summary>
    /// Views details of a specific calculated column
    /// </summary>
    int ViewCalculatedColumn(string[] args);

    /// <summary>
    /// Updates an existing calculated column
    /// </summary>
    int UpdateCalculatedColumn(string[] args);

    /// <summary>
    /// Deletes a calculated column from a table
    /// </summary>
    int DeleteCalculatedColumn(string[] args);

    /// <summary>
    /// Validates a DAX formula without creating/updating objects
    /// </summary>
    int ValidateDax(string[] args);
}

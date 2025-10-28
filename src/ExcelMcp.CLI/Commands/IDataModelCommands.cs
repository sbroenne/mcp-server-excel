namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Interface for Data Model CLI commands
/// </summary>
public interface IDataModelCommands
{
    /// <summary>
    /// Lists all tables in the Data Model
    /// Usage: dm-list-tables <file.xlsx>
    /// </summary>
    int ListTables(string[] args);

    /// <summary>
    /// Lists all DAX measures in the Data Model
    /// Usage: dm-list-measures <file.xlsx>
    /// </summary>
    int ListMeasures(string[] args);

    /// <summary>
    /// Views a specific DAX measure formula
    /// Usage: dm-view-measure <file.xlsx> <measure-name>
    /// </summary>
    int ViewMeasure(string[] args);

    /// <summary>
    /// Exports a DAX measure to a file
    /// Usage: dm-export-measure <file.xlsx> <measure-name> <output.dax>
    /// </summary>
    int ExportMeasure(string[] args);

    /// <summary>
    /// Lists all relationships in the Data Model
    /// Usage: dm-list-relationships <file.xlsx>
    /// </summary>
    int ListRelationships(string[] args);

    /// <summary>
    /// Refreshes the Data Model
    /// Usage: dm-refresh <file.xlsx>
    /// </summary>
    int Refresh(string[] args);

    /// <summary>
    /// Deletes a DAX measure from the Data Model
    /// Usage: dm-delete-measure <file.xlsx> <measure-name>
    /// </summary>
    int DeleteMeasure(string[] args);

    /// <summary>
    /// Deletes a relationship from the Data Model
    /// Usage: dm-delete-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column>
    /// </summary>
    int DeleteRelationship(string[] args);

    // Phase 2: Discovery operations

    /// <summary>
    /// Lists all columns in a Data Model table
    /// Usage: dm-list-columns <file.xlsx> <table-name>
    /// </summary>
    int ListColumns(string[] args);

    /// <summary>
    /// Views detailed information about a Data Model table
    /// Usage: dm-view-table <file.xlsx> <table-name>
    /// </summary>
    int ViewTable(string[] args);

    /// <summary>
    /// Gets Data Model overview (table/measure/relationship counts)
    /// Usage: dm-get-model-info <file.xlsx>
    /// </summary>
    int GetModelInfo(string[] args);

    // Phase 2: CREATE/UPDATE operations

    /// <summary>
    /// Creates a new DAX measure in the Data Model
    /// Usage: dm-create-measure <file.xlsx> <table-name> <measure-name> <dax-formula> [format-type] [description]
    /// </summary>
    int CreateMeasure(string[] args);

    /// <summary>
    /// Updates an existing DAX measure
    /// Usage: dm-update-measure <file.xlsx> <measure-name> [dax-formula] [format-type] [description]
    /// </summary>
    int UpdateMeasure(string[] args);

    /// <summary>
    /// Creates a relationship between two Data Model tables
    /// Usage: dm-create-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> [active:true|false]
    /// </summary>
    int CreateRelationship(string[] args);

    /// <summary>
    /// Updates a relationship's active status
    /// Usage: dm-update-relationship <file.xlsx> <from-table> <from-column> <to-table> <to-column> <active:true|false>
    /// </summary>
    int UpdateRelationship(string[] args);
}

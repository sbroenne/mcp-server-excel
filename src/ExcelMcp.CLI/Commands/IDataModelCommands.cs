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
    Task<int> ExportMeasure(string[] args);

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
}

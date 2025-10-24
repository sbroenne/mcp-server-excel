namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Table management commands for CLI
/// </summary>
public interface ITableCommands
{
    /// <summary>
    /// Lists all Excel Tables in a workbook
    /// </summary>
    int List(string[] args);

    /// <summary>
    /// Creates a new Excel Table
    /// </summary>
    int Create(string[] args);

    /// <summary>
    /// Renames an Excel Table
    /// </summary>
    int Rename(string[] args);

    /// <summary>
    /// Deletes an Excel Table (converts to range)
    /// </summary>
    int Delete(string[] args);

    /// <summary>
    /// Gets detailed information about an Excel Table
    /// </summary>
    int Info(string[] args);
}

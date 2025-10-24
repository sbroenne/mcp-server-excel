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

    /// <summary>
    /// Resizes an Excel Table
    /// </summary>
    int Resize(string[] args);

    /// <summary>
    /// Toggles the totals row
    /// </summary>
    int ToggleTotals(string[] args);

    /// <summary>
    /// Sets column total function
    /// </summary>
    int SetColumnTotal(string[] args);

    /// <summary>
    /// Reads data from a table
    /// </summary>
    int ReadData(string[] args);

    /// <summary>
    /// Appends rows to a table
    /// </summary>
    int AppendRows(string[] args);

    /// <summary>
    /// Changes table style
    /// </summary>
    int SetStyle(string[] args);

    /// <summary>
    /// Adds table to Power Pivot Data Model
    /// </summary>
    int AddToDataModel(string[] args);
}

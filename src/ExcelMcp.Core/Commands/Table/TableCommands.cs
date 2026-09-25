namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Excel Table (ListObject) management commands - main partial class with shared state and helper methods
/// </summary>
public partial class TableCommands : ITableCommands, ITableColumnCommands
{
    #region Helper Methods

    /// <summary>
    /// Finds a table by name in the workbook, throwing if not found.
    /// Delegates to CoreLookupHelpers.FindTable for the actual lookup.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to find</param>
    /// <returns>The table object (caller must release)</returns>
    /// <exception cref="InvalidOperationException">Thrown if table is not found</exception>
    private static dynamic FindTable(dynamic workbook, string tableName)
        => CoreLookupHelpers.FindTable(workbook, tableName);

    /// <summary>
    /// Checks if a table with the given name exists in the workbook
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to check</param>
    /// <returns>True if table exists, false otherwise</returns>
    private static bool TableExists(dynamic workbook, string tableName)
        => CoreLookupHelpers.TableExists(workbook, tableName);

    #endregion
}

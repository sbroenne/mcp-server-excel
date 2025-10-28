using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Excel Table (ListObject) management commands - main partial class with shared state and helper methods
/// </summary>
public partial class TableCommands : ITableCommands
{
    #region Constants and Validation

    /// <summary>
    /// Regex pattern for valid table names
    /// </summary>
    private static readonly Regex TableNameRegex = new(@"^[a-zA-Z_][a-zA-Z0-9_]*$", RegexOptions.Compiled);

    /// <summary>
    /// Maximum allowed table name length
    /// </summary>
    private const int MaxTableNameLength = 255;

    /// <summary>
    /// Validates table name format
    /// </summary>
    private static bool IsValidTableName(string name) => TableNameRegex.IsMatch(name);

    /// <summary>
    /// Validates a table name to prevent injection attacks and ensure Excel compatibility
    /// </summary>
    /// <param name="tableName">Table name to validate</param>
    /// <exception cref="ArgumentException">Thrown if table name is invalid</exception>
    private static void ValidateTableName(string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Table name cannot be null or empty", nameof(tableName));
        }

        if (tableName.Length > MaxTableNameLength)
        {
            throw new ArgumentException(
                $"Table name too long: {tableName.Length} characters (maximum: {MaxTableNameLength})",
                nameof(tableName));
        }

        if (!TableNameRegex.IsMatch(tableName))
        {
            throw new ArgumentException(
                $"Invalid table name '{tableName}'. Table names must start with a letter or underscore, " +
                "and can only contain letters, numbers, and underscores (no spaces or special characters).",
                nameof(tableName));
        }

        // Check for reserved names
        string upperName = tableName.ToUpperInvariant();
        if (upperName == "PRINT_AREA" || upperName == "PRINT_TITLES" ||
            upperName == "_XLNM" || upperName.StartsWith("_XLNM."))
        {
            throw new ArgumentException(
                $"Table name '{tableName}' is reserved by Excel",
                nameof(tableName));
        }
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Finds a table by name in the workbook
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to find</param>
    /// <returns>The table object if found, null otherwise</returns>
    private static dynamic? FindTable(dynamic workbook, string tableName)
    {
        dynamic? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic? sheet = null;
                dynamic? listObjects = null;
                try
                {
                    sheet = sheets.Item(i);
                    listObjects = sheet.ListObjects;

                    for (int j = 1; j <= listObjects.Count; j++)
                    {
                        dynamic? table = null;
                        try
                        {
                            table = listObjects.Item(j);
                            if (table.Name == tableName)
                            {
                                // Found it - return without releasing
                                return table;
                            }
                        }
                        finally
                        {
                            if (table != null && table.Name != tableName)
                            {
                                // Only release if not returning this table
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
                    ComUtilities.Release(ref sheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }

        return null;
    }

    /// <summary>
    /// Checks if a table with the given name exists in the workbook
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to check</param>
    /// <returns>True if table exists, false otherwise</returns>
    private static bool TableExists(dynamic workbook, string tableName)
    {
        dynamic? table = FindTable(workbook, tableName);
        bool exists = table != null;
        ComUtilities.Release(ref table);
        return exists;
    }

    #endregion
}

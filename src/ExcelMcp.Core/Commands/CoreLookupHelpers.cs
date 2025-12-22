using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Shared lookup helpers for finding Excel objects (PivotTables, Tables, etc.) across the workbook.
/// These utilities centralize common lookup patterns to avoid code duplication.
/// </summary>
public static class CoreLookupHelpers
{
    #region PivotTable Lookup

    /// <summary>
    /// Tries to find a PivotTable by name across all worksheets.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="pivotTableName">Name of the PivotTable to find</param>
    /// <param name="pivotTable">The found PivotTable object (caller must release), or null if not found</param>
    /// <returns>True if found, false otherwise</returns>
    /// <remarks>
    /// Caller is responsible for releasing the returned COM object using ComUtilities.Release().
    /// </remarks>
    public static bool TryFindPivotTable(dynamic workbook, string pivotTableName, out dynamic? pivotTable)
    {
        pivotTable = null;
        dynamic? sheets = null;

        try
        {
            sheets = workbook.Worksheets;
            int sheetCount = Convert.ToInt32(sheets.Count);

            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic? sheet = null;
                dynamic? pivotTables = null;

                try
                {
                    sheet = sheets.Item(i);
                    pivotTables = sheet.PivotTables();
                    int ptCount = Convert.ToInt32(pivotTables.Count);

                    for (int j = 1; j <= ptCount; j++)
                    {
                        dynamic? pt = null;

                        try
                        {
                            pt = pivotTables.Item(j);
                            string ptName = pt.Name?.ToString() ?? string.Empty;

                            if (ptName.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                            {
                                // Found - release intermediate objects but NOT the found PivotTable
                                ComUtilities.Release(ref pivotTables!);
                                ComUtilities.Release(ref sheet!);
                                ComUtilities.Release(ref sheets!);
                                pivotTable = pt;
                                return true;
                            }
                        }
                        finally
                        {
                            // Only release if not the found item
                            if (pt != null && pivotTable == null)
                            {
                                ComUtilities.Release(ref pt!);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivotTables);
                    ComUtilities.Release(ref sheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }

        return false;
    }

    /// <summary>
    /// Finds a PivotTable by name across all worksheets, throwing if not found.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="pivotTableName">Name of the PivotTable to find</param>
    /// <returns>The PivotTable object (caller must release)</returns>
    /// <exception cref="InvalidOperationException">Thrown if PivotTable is not found</exception>
    /// <remarks>
    /// Caller is responsible for releasing the returned COM object using ComUtilities.Release().
    /// </remarks>
    public static dynamic FindPivotTable(dynamic workbook, string pivotTableName)
    {
        if (!TryFindPivotTable(workbook, pivotTableName, out dynamic? pivotTable) || pivotTable == null)
        {
            throw new InvalidOperationException($"PivotTable '{pivotTableName}' not found.");
        }

        return pivotTable!;
    }

    #endregion

    #region Table Lookup

    /// <summary>
    /// Tries to find an Excel Table (ListObject) by name in the workbook.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to find</param>
    /// <param name="table">The found table object (caller must release), or null if not found</param>
    /// <returns>True if found, false otherwise</returns>
    /// <remarks>
    /// Caller is responsible for releasing the returned COM object using ComUtilities.Release().
    /// </remarks>
    public static bool TryFindTable(dynamic workbook, string tableName, out dynamic? table)
    {
        table = null;
        dynamic? sheets = null;

        try
        {
            sheets = workbook.Worksheets;
            int sheetCount = Convert.ToInt32(sheets.Count);

            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic? sheet = null;
                dynamic? listObjects = null;

                try
                {
                    sheet = sheets.Item(i);
                    listObjects = sheet.ListObjects;
                    int tableCount = Convert.ToInt32(listObjects.Count);

                    for (int j = 1; j <= tableCount; j++)
                    {
                        dynamic? tbl = null;

                        try
                        {
                            tbl = listObjects.Item(j);
                            string tblName = tbl.Name?.ToString() ?? string.Empty;

                            if (tblName.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                            {
                                // Found - release intermediate objects but NOT the found table
                                ComUtilities.Release(ref listObjects!);
                                ComUtilities.Release(ref sheet!);
                                ComUtilities.Release(ref sheets!);
                                table = tbl;
                                return true;
                            }
                        }
                        finally
                        {
                            // Only release if not the found item
                            if (tbl != null && table == null)
                            {
                                ComUtilities.Release(ref tbl!);
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

        return false;
    }

    /// <summary>
    /// Finds an Excel Table (ListObject) by name in the workbook, throwing if not found.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to find</param>
    /// <returns>The table object (caller must release)</returns>
    /// <exception cref="InvalidOperationException">Thrown if table is not found</exception>
    /// <remarks>
    /// Caller is responsible for releasing the returned COM object using ComUtilities.Release().
    /// </remarks>
    public static dynamic FindTable(dynamic workbook, string tableName)
    {
        if (!TryFindTable(workbook, tableName, out dynamic? table) || table == null)
        {
            throw new InvalidOperationException($"Table '{tableName}' not found.");
        }

        return table!;
    }

    /// <summary>
    /// Checks if a table with the given name exists in the workbook.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to check</param>
    /// <returns>True if table exists, false otherwise</returns>
    public static bool TableExists(dynamic workbook, string tableName)
    {
        if (TryFindTable(workbook, tableName, out dynamic? table))
        {
            ComUtilities.Release(ref table);
            return true;
        }

        return false;
    }

    #endregion
}

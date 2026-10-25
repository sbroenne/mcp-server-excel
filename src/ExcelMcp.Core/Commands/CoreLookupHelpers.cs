using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

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
    public static bool TryFindPivotTable(Excel.Workbook workbook, string pivotTableName, out Excel.PivotTable? pivotTable)
    {
        pivotTable = null;
        Excel.Sheets? sheets = null;

        try
        {
            sheets = workbook.Worksheets;
            int sheetCount = sheets.Count;

            for (int i = 1; i <= sheetCount; i++)
            {
                Excel.Worksheet? sheet = null;
                Excel.PivotTables? pivotTables = null;

                try
                {
                    sheet = (Excel.Worksheet)sheets[i];
                    pivotTables = (Excel.PivotTables)sheet.PivotTables();
                    int ptCount = pivotTables.Count;

                    for (int j = 1; j <= ptCount; j++)
                    {
                        Excel.PivotTable? pt = null;

                        try
                        {
                            pt = pivotTables.Item(j);
                            string ptName = pt.Name;

                            if (ptName.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                            {
                                pivotTable = pt;
                                pt = null;
                                return true;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref pt);
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
    public static Excel.PivotTable FindPivotTable(Excel.Workbook workbook, string pivotTableName)
    {
        if (!TryFindPivotTable(workbook, pivotTableName, out Excel.PivotTable? pivotTable) || pivotTable == null)
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
    public static bool TryFindTable(Excel.Workbook workbook, string tableName, out Excel.ListObject? table)
    {
        table = null;
        Excel.Sheets? sheets = null;

        try
        {
            sheets = workbook.Worksheets;
            int sheetCount = sheets.Count;

            for (int i = 1; i <= sheetCount; i++)
            {
                Excel.Worksheet? sheet = null;
                Excel.ListObjects? listObjects = null;

                try
                {
                    sheet = (Excel.Worksheet)sheets[i];
                    listObjects = sheet.ListObjects;
                    int tableCount = listObjects.Count;

                    for (int j = 1; j <= tableCount; j++)
                    {
                        Excel.ListObject? tbl = null;

                        try
                        {
                            tbl = listObjects[j];
                            string tblName = tbl.Name;

                            if (tblName.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                            {
                                table = tbl;
                                tbl = null;
                                return true;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref tbl);
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
    public static Excel.ListObject FindTable(Excel.Workbook workbook, string tableName)
    {
        if (!TryFindTable(workbook, tableName, out Excel.ListObject? table) || table == null)
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
    public static bool TableExists(Excel.Workbook workbook, string tableName)
    {
        if (TryFindTable(workbook, tableName, out Excel.ListObject? table))
        {
            ComUtilities.Release(ref table);
            return true;
        }

        return false;
    }

    #endregion
}


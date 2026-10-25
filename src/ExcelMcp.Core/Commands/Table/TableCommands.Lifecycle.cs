using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table lifecycle operations (List, Create, Rename, Delete, GetInfo)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public TableListResult List(IExcelBatch batch)
    {
        var result = new TableListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    Excel.Worksheet? sheet = null;
                    Excel.ListObjects? listObjects = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[i];
                        listObjects = sheet.ListObjects;
                        string sheetName = sheet.Name;

                        for (int j = 1; j <= listObjects.Count; j++)
                        {
                            Excel.ListObject? table = null;
                            Excel.Range? dataBodyRange = null;
                            Excel.Range? tableRange = null;
                            Excel.ListColumns? listColumns = null;
                            Excel.Range? dataRows = null;
                            try
                            {
                                table = listObjects[j];
                                string tableName = table.Name;
                                tableRange = table.Range;
                                string rangeAddress = tableRange.Address;
                                bool showHeaders = table.ShowHeaders;
                                bool showTotals = table.ShowTotals;
                                string tableStyleName = GetTableStyleName(table);

                                // Get column count and names
                                listColumns = table.ListColumns;
                                int columnCount = listColumns.Count;
                                var columns = new List<string>();

                                if (showHeaders)
                                {
                                    for (int k = 1; k <= listColumns.Count; k++)
                                    {
                                        Excel.ListColumn? column = null;
                                        try
                                        {
                                            column = listColumns[k];
                                            columns.Add(column.Name);
                                        }
                                        finally
                                        {
                                            ComUtilities.Release(ref column);
                                        }
                                    }
                                }

                                // Get row count (excluding header)
                                // SECURITY FIX: DataBodyRange can be NULL if table has only headers
                                int rowCount = 0;
                                try
                                {
                                    dataBodyRange = table.DataBodyRange;
                                    if (dataBodyRange != null)
                                    {
                                        dataRows = dataBodyRange.Rows;
                                        rowCount = dataRows.Count;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref dataRows);
                                    ComUtilities.Release(ref dataBodyRange);
                                }

                                result.Tables.Add(new TableInfo
                                {
                                    Name = tableName,
                                    SheetName = sheetName,
                                    Range = rangeAddress,
                                    HasHeaders = showHeaders,
                                    TableStyle = tableStyleName,
                                    RowCount = rowCount,
                                    ColumnCount = columnCount,
                                    Columns = columns,
                                    ShowTotals = showTotals
                                });
                            }
                            finally
                            {
                                ComUtilities.Release(ref listColumns);
                                ComUtilities.Release(ref tableRange);
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref listObjects);
                        ComUtilities.Release(ref sheet);
                    }
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Create(IExcelBatch batch, string sheetName, string tableName, string rangeAddress, bool hasHeaders = true, string? tableStyle = null)
    {
        ValidateCreateInputs(sheetName, tableName, rangeAddress);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? rangeObj = null;
            Excel.ListObjects? listObjects = null;
            Excel.ListObject? newTable = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                rangeObj = ResolveEffectiveRange(sheet, rangeAddress);
                TablePreflightResult preflight = AnalyzePreflight(
                    ctx.Book,
                    rangeObj,
                    batch.WorkbookPath,
                    sheetName,
                    tableName,
                    rangeAddress,
                    hasHeaders,
                    ct);
                if (!preflight.SafeToCreate)
                {
                    string reasons = string.Join(
                        " ",
                        preflight.Findings
                            .Where(finding => finding.Severity == TablePreflightSeverity.Blocker)
                            .Select(finding => finding.Message));
                    throw new InvalidOperationException(
                        $"Table '{tableName}' cannot be created because preflight found blocking issues. {reasons}");
                }

                listObjects = sheet.ListObjects;

                var headerOption = hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlGuess;
                newTable = listObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rangeObj, null, headerOption);
                newTable.Name = tableName;

                // Apply table style if specified
                if (!string.IsNullOrWhiteSpace(tableStyle))
                {
                    newTable.TableStyle = tableStyle;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref newTable);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref rangeObj);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Rename(IExcelBatch batch, string tableName, string newName)
    {
        ValidateRequiredTableName(tableName);
        ValidateRequiredTableName(newName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Check if new name already exists
                if (TableExists(ctx.Book, newName))
                {
                    throw new InvalidOperationException($"Table '{newName}' already exists");
                }

                table.Name = newName;
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string tableName)
    {
        ValidateRequiredTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Convert table back to range (Unlist)
                table.Unlist();

                // SECURITY FIX: After Unlist(), we must explicitly release the table COM object
                // The table object is no longer valid but still holds a COM reference
                ComUtilities.Release(ref table);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public TableInfoResult Read(IExcelBatch batch, string tableName)
    {
        ValidateRequiredTableName(tableName);

        var result = new TableInfoResult { FilePath = batch.WorkbookPath };
        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? dataBodyRange = null;
            Excel.Range? tableRange = null;
            Excel.ListColumns? listColumns = null;
            Excel.Range? dataRows = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                sheet = (Excel.Worksheet)(object)table.Parent;
                string sheetName = sheet.Name;
                tableRange = table.Range;
                string rangeAddress = tableRange.Address;
                bool showHeaders = table.ShowHeaders;
                bool showTotals = table.ShowTotals;
                string tableStyleName = GetTableStyleName(table);

                // Get column count and names
                listColumns = table.ListColumns;
                int columnCount = listColumns.Count;
                var columns = new List<string>();

                if (showHeaders)
                {
                    for (int i = 1; i <= listColumns.Count; i++)
                    {
                        Excel.ListColumn? column = null;
                        try
                        {
                            column = listColumns[i];
                            columns.Add(column.Name);
                        }
                        finally
                        {
                            ComUtilities.Release(ref column);
                        }
                    }
                }

                // Get row count (excluding header)
                // SECURITY FIX: DataBodyRange can be NULL if table has only headers
                int rowCount = 0;
                try
                {
                    dataBodyRange = table.DataBodyRange;
                    if (dataBodyRange != null)
                    {
                        dataRows = dataBodyRange.Rows;
                        rowCount = dataRows.Count;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref dataRows);
                    ComUtilities.Release(ref dataBodyRange);
                }

                result.Table = new TableInfo
                {
                    Name = tableName,
                    SheetName = sheetName,
                    Range = rangeAddress,
                    HasHeaders = showHeaders,
                    TableStyle = tableStyleName,
                    RowCount = rowCount,
                    ColumnCount = columnCount,
                    Columns = columns,
                    ShowTotals = showTotals
                };

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }
    private static string GetTableStyleName(Excel.ListObject table)
    {
        object? tableStyle = null;
        try
        {
            tableStyle = table.TableStyle;
            return tableStyle switch
            {
                // Excel can represent a cleared style as a single NUL character.
                null or "\0" => "",
                string styleName => styleName,
                _ => ((Excel.TableStyle)tableStyle).Name ?? ""
            };
        }
        finally
        {
            ComUtilities.Release(ref tableStyle);
        }
    }
}

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class PowerQueryCommands
{
    private static void ReadLoadMetrics(
        Excel.Workbook workbook, string queryName, PowerQueryOperationResult result, CancellationToken ct)
    {
        Excel.Sheets? sheets = null;
        Excel.Model? model = null;
        Excel.ModelTables? modelTables = null;
        try
        {
            sheets = workbook.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Worksheet? sheet = null;
                Excel.ListObjects? tables = null;
                try
                {
                    sheet = (Excel.Worksheet)sheets.Item[i];
                    tables = sheet.ListObjects;
                    for (int j = 1; j <= tables.Count; j++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.ListObject? table = null;
                        Excel.QueryTable? queryTable = null;
                        Excel.ListRows? rows = null;
                        try
                        {
                            table = tables.Item[j];
                            if (table.SourceType is not (Excel.XlListObjectSourceType.xlSrcExternal or Excel.XlListObjectSourceType.xlSrcQuery))
                                continue;
                            queryTable = table.QueryTable;
                            if (queryTable == null || !MatchesWorksheetQueryTable(queryTable, queryName))
                                continue;
                            rows = table.ListRows;
                            result.WorksheetRowsLoaded = (result.WorksheetRowsLoaded ?? 0) + rows.Count;
                        }
                        finally
                        {
                            ComUtilities.Release(ref rows);
                            ComUtilities.Release(ref queryTable);
                            ComUtilities.Release(ref table);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref tables);
                    ComUtilities.Release(ref sheet);
                }
            }

            model = workbook.Model;
            modelTables = model.ModelTables;
            for (int i = 1; i <= modelTables.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.ModelTable? table = null;
                Excel.WorkbookConnection? connection = null;
                try
                {
                    table = modelTables.Item(i);
                    connection = table.SourceWorkbookConnection;
                    if (connection != null && MatchesDataModelSourceConnection(connection, queryName))
                        result.ModelRowsLoaded = (result.ModelRowsLoaded ?? 0) + table.RecordCount;
                }
                finally
                {
                    ComUtilities.Release(ref connection);
                    ComUtilities.Release(ref table);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref sheets);
        }

        result.RowsLoaded = (result.WorksheetRowsLoaded, result.ModelRowsLoaded) switch
        {
            ({ } worksheet, { } dataModel) => worksheet == dataModel ? worksheet : null,
            ({ } worksheet, null) => worksheet,
            (null, { } dataModel) => dataModel,
            _ => null
        };
    }
}

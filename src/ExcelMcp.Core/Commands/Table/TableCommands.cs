using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Excel Table (ListObject) management commands - main partial class with shared state and helper methods
/// </summary>
public partial class TableCommands : ITableCommands, ITableColumnCommands
{
    #region Helper Methods

    private static void ValidateRequiredTableName(string tableName)
        => ArgumentException.ThrowIfNullOrWhiteSpace(tableName);

    private static bool ValidateTableNameWithExcel(Excel.Application application, Excel.Workbook sourceWorkbook, string tableName)
    {
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? scratchWorkbook = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? range = null;
        Excel.ListObjects? tables = null;
        Excel.ListObject? table = null;
        try
        {
            // Probe Excel's own naming rules without converting any source cells to a table.
            workbooks = application.Workbooks;
            scratchWorkbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            sheets = scratchWorkbook.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            range = sheet.Range["A1:A2"];
            range.Value2 = new object[,] { { "Header" }, { 1 } };
            tables = sheet.ListObjects;
            table = tables.Add(Excel.XlListObjectSourceType.xlSrcRange, range,
                Type.Missing, Excel.XlYesNoGuess.xlYes);
            return TryAssignTableNameForValidation(table, tableName);
        }
        finally
        {
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref tables);
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
            try
            {
                if (scratchWorkbook != null)
                {
                    try
                    {
                        ExcelShutdownService.CloseWorkbookOrThrow(scratchWorkbook, save: false);
                    }
                    finally
                    {
                        ComUtilities.Release(ref scratchWorkbook);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref workbooks);
                sourceWorkbook.Activate();
            }
        }
    }

    internal static bool TryAssignTableNameForValidation(dynamic table, string tableName)
    {
        try
        {
            table.Name = tableName;
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (COMException ex) when (!IsExcelSessionFailure(ex.HResult))
        {
            return false;
        }
    }

    private static bool IsExcelSessionFailure(int hResult)
        => hResult is
            ResiliencePipelines.RPC_E_SERVERCALL_RETRYLATER or
            ResiliencePipelines.RPC_E_CALL_REJECTED or
            ResiliencePipelines.RPC_E_CALL_FAILED or
            ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE or
            ResiliencePipelines.RPC_E_DISCONNECTED or
            ResiliencePipelines.CO_E_SERVER_EXEC_FAILURE or
            ResiliencePipelines.DATA_MODEL_BUSY;

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

    internal static void SetCreatedTableNameOrRollback(
        dynamic listObject,
        string tableName,
        dynamic? workbookConnection = null,
        bool preserveSourceRange = false)
    {
        try
        {
            listObject.Name = tableName;
        }
        catch (Exception nameException)
        {
            var cleanupErrors = new List<Exception>();
            TryRemoveCreatedTable(
                listObject,
                $"the default-named table created before Excel rejected '{tableName}'",
                preserveSourceRange,
                cleanupErrors);

            if (workbookConnection != null)
            {
                TryDeleteCreatedObject(
                    workbookConnection,
                    $"the workbook connection created before Excel rejected '{tableName}'",
                    cleanupErrors);
            }

            if (cleanupErrors.Count > 0)
            {
                cleanupErrors.Insert(0, nameException);
                throw new InvalidOperationException(
                    $"Excel rejected table name '{tableName}', and cleanup of created workbook objects failed.",
                    new AggregateException(cleanupErrors));
            }

            throw;
        }
    }

    private static void TryRemoveCreatedTable(
        dynamic value,
        string description,
        bool preserveSourceRange,
        List<Exception> cleanupErrors)
    {
        try
        {
            if (preserveSourceRange)
            {
                value.Unlist();
            }
            else
            {
                value.Delete();
            }
        }
        catch (COMException ex)
        {
            cleanupErrors.Add(new InvalidOperationException($"Failed to remove {description}.", ex));
        }
    }

    private static void TryDeleteCreatedObject(dynamic value, string description, List<Exception> cleanupErrors)
    {
        try
        {
            value.Delete();
        }
        catch (COMException ex)
        {
            cleanupErrors.Add(new InvalidOperationException($"Failed to delete {description}.", ex));
        }
    }

    #endregion
}

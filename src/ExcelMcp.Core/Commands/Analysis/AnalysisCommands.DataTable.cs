using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

public sealed partial class AnalysisCommands
{
    /// <inheritdoc />
    public OperationResult CreateDataTable(
        IExcelBatch batch,
        string sheetName,
        string tableRange,
        string? rowInputCell = null,
        string? columnInputCell = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(tableRange);
        if (string.IsNullOrWhiteSpace(rowInputCell) && string.IsNullOrWhiteSpace(columnInputCell))
        {
            throw new ArgumentException("Provide rowInputCell, columnInputCell, or both for a data table.");
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? table = null;
            Excel.Range? rowInput = null;
            Excel.Range? columnInput = null;
            object? tableResult = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                table = sheet.Range[tableRange];
                if (!string.IsNullOrWhiteSpace(rowInputCell))
                {
                    rowInput = sheet.Range[rowInputCell];
                }
                if (!string.IsNullOrWhiteSpace(columnInputCell))
                {
                    columnInput = sheet.Range[columnInputCell];
                }

                tableResult = table.Table(
                    (object?)rowInput ?? Type.Missing,
                    (object?)columnInput ?? Type.Missing);
                return new OperationResult
                {
                    Success = true,
                    Message = $"Data table created in '{sheetName}'!{tableRange}."
                };
            }
            finally
            {
                ComUtilities.Release(ref tableResult);
                ComUtilities.Release(ref columnInput);
                ComUtilities.Release(ref rowInput);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

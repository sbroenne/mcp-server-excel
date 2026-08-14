using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable drill-through operations.
/// </summary>
public partial class PivotTableCommands
{
    /// <inheritdoc />
    public PivotDrillThroughResult DrillThrough(
        IExcelBatch batch,
        string pivotTableName,
        string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.Worksheet? pivotSheet = null;
            Excel.Range? targetCell = null;
            Excel.Range? dataBodyRange = null;
            Excel.Range? intersection = null;
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? detailSheet = null;
            Excel.Range? usedRange = null;
            Excel.Range? usedRows = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                cache = pivot.PivotCache();
                if (cache.OLAP)
                {
                    throw new InvalidOperationException(
                        "OLAP/Data Model drill-through is provider-dependent and is not supported by this action. " +
                        "Use a regular PivotTable source for deterministic detail extraction.");
                }

                pivotSheet = (Excel.Worksheet)pivot.Parent;
                targetCell = pivotSheet.Range[cellAddress];
                dataBodyRange = pivot.DataBodyRange;
                intersection = ctx.App.Intersect(targetCell, dataBodyRange);
                if (intersection == null)
                {
                    throw new ArgumentException(
                        $"Cell '{cellAddress}' is not inside the data body of PivotTable '{pivotTableName}'.",
                        nameof(cellAddress));
                }

                worksheets = ctx.Book.Worksheets;
                var existingSheetNames = new HashSet<string>(
                    Enumerable.Range(1, worksheets.Count)
                        .Select(index =>
                        {
                            Excel.Worksheet? sheet = null;
                            try
                            {
                                sheet = (Excel.Worksheet)worksheets[index];
                                return sheet.Name;
                            }
                            finally
                            {
                                ComUtilities.Release(ref sheet);
                            }
                        }),
                    StringComparer.OrdinalIgnoreCase);

                targetCell.ShowDetail = true;

                for (var index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet? candidate = null;
                    try
                    {
                        candidate = (Excel.Worksheet)worksheets[index];
                        if (!existingSheetNames.Contains(candidate.Name))
                        {
                            detailSheet = candidate;
                            candidate = null;
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref candidate);
                    }
                }

                if (detailSheet == null)
                {
                    throw new InvalidOperationException(
                        $"Excel did not create a detail worksheet for PivotTable cell '{cellAddress}'.");
                }

                usedRange = detailSheet.UsedRange;
                usedRows = usedRange.Rows;
                return new PivotDrillThroughResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    CellAddress = targetCell.Address,
                    DetailSheetName = detailSheet.Name,
                    DetailRowCount = usedRows.Count,
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref usedRows);
                ComUtilities.Release(ref usedRange);
                ComUtilities.Release(ref detailSheet);
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref intersection);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref targetCell);
                ComUtilities.Release(ref pivotSheet);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
            }
        });
    }
}

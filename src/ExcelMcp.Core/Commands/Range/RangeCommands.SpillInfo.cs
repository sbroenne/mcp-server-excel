using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeSpillInfoResult GetSpillInfo(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        return batch.Execute((ctx, ct) =>
        {
            if (!ctx.Capabilities.SupportsFormula2)
            {
                throw new NotSupportedException(
                    "Native spill inspection is unavailable: this Excel session does not support dynamic-array Formula2.");
            }
            Excel.Range? range = null;
            Excel.Worksheet? sheet = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                sheet = (Excel.Worksheet)range!.Parent;
                var result = new RangeSpillInfoResult
                {
                    FilePath = batch.WorkbookPath,
                    SheetName = sheet.Name,
                    RangeAddress = range.Address
                };
                RangeHelpers.VisitCells(range, ct, cell => result.Cells.Add(ReadSpillCell(cell)));
                result.Cells = result.Cells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToList();
                result.CellCount = result.Cells.Count;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref range);
            }
        });
    }

    private static SpillCellInfo ReadSpillCell(Excel.Range cell)
    {
        string address = cell.Address;
        int row = cell.Row;
        int column = cell.Column;
        object hasSpillValue = cell.HasSpill;
        if (hasSpillValue is not bool hasSpill)
        {
            throw new InvalidOperationException($"Excel returned an indeterminate spill state for {address}.");
        }
        if (!hasSpill)
        {
            bool hasFormula = Convert.ToBoolean((object)cell.HasFormula, CultureInfo.InvariantCulture);
            object? value = cell.Value2;
            if (hasFormula && ExcelErrorMapper.TryGet(value, out _, out var error) && error.Name == "#SPILL!")
            {
                return new SpillCellInfo(address, row, column, SpillCellState.Blocked,
                    address, Convert.ToString((object)cell.Formula2, CultureInfo.InvariantCulture));
            }
            return new SpillCellInfo(address, row, column, SpillCellState.Ordinary);
        }
        Excel.Range? source = null;
        Excel.Range? spill = null;
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            source = cell.SpillParent;
            spill = source.SpillingToRange;
            rows = spill.Rows;
            columns = spill.Columns;
            return new SpillCellInfo(address, row, column,
                source.Row == row && source.Column == column ? SpillCellState.Source : SpillCellState.Result,
                source.Address, Convert.ToString((object)source.Formula2, CultureInfo.InvariantCulture),
                spill.Address, rows.Count, columns.Count);
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref spill);
            ComUtilities.Release(ref source);
        }
    }
}

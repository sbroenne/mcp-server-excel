using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult SetCellProtection(IExcelBatch batch, string sheetName,
        string rangeAddress, bool? locked = null, bool? formulaHidden = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        if (!locked.HasValue && !formulaHidden.HasValue)
            throw new ArgumentException("Supply locked, formulaHidden, or both.");
        return batch.Execute((context, ct) =>
        {
            Excel.Range? scope = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                scope = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (scope is null)
                    throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                if (locked.HasValue) scope.Locked = locked.Value;
                if (formulaHidden.HasValue) scope.FormulaHidden = formulaHidden.Value;
                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "set-cell-protection",
                    Message = "Supplied native protection flags changed; enforcement requires sheet protection."
                };
            }
            finally
            {
                ComUtilities.Release(ref scope);
            }
        });
    }

    /// <inheritdoc />
    public RangeCellProtectionResult GetCellProtection(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        return batch.Execute((context, ct) =>
        {
            Excel.Range? scope = null;
            Excel.Worksheet? sheet = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                scope = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (scope is null)
                    throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                sheet = scope.Worksheet;
                var result = new RangeCellProtectionResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "get-cell-protection",
                    SheetName = sheet.Name,
                    RangeAddress = scope.Address
                };
                RangeHelpers.VisitCells(scope, ct, cell =>
                {
                    if (cell.Locked is not bool locked || cell.FormulaHidden is not bool formulaHidden)
                        throw new InvalidOperationException($"Excel returned indeterminate protection for cell '{cell.Address}'.");
                    result.Cells.Add(new CellProtectionRead(cell.Address, locked, formulaHidden));
                });
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref scope);
            }
        });
    }
}

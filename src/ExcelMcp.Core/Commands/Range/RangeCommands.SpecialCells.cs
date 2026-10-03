using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public SpecialCellsResult GetSpecialCells(
        IExcelBatch batch, string sheetName, string rangeAddress, SpecialCellKind cellKind)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        if (!Enum.IsDefined(cellKind))
        {
            throw new ArgumentOutOfRangeException(nameof(cellKind), cellKind, "Unknown cell selector.");
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Worksheet? sheet = null;
            Excel.Areas? areas = null;
            Excel.WorksheetFunction? functions = null;
            Excel.Range? usedRange = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range is null)
                {
                    throw new InvalidOperationException(RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                sheet = (Excel.Worksheet)range.Parent;
                if (cellKind == SpecialCellKind.Errors)
                {
                    functions = ctx.App.WorksheetFunction;
                }
                if (cellKind == SpecialCellKind.Blanks)
                {
                    usedRange = sheet.UsedRange;
                }

                var result = new SpecialCellsResult
                {
                    FilePath = batch.WorkbookPath,
                    SheetName = sheet.Name,
                    RangeAddress = range.Address,
                    CellKind = cellKind
                };
                var matches = new List<SpecialCellArea>();
                areas = range.Areas;
                int areaCount = areas.Count;
                for (int index = 1; index <= areaCount; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? area = null;
                    try
                    {
                        area = areas[index];
                        // SpecialCells expands a one-cell input to UsedRange.
                        if (Convert.ToInt64(area.CountLarge, CultureInfo.InvariantCulture) == 1)
                        {
                            if (MatchesCell(area, cellKind, functions))
                            {
                                AddSpecialCellArea(area, matches);
                            }
                            continue;
                        }

                        if (cellKind == SpecialCellKind.Blanks)
                        {
                            CollectBlankCells(area, usedRange!, sheet, ctx.App, matches, ct);
                        }
                        else if (cellKind == SpecialCellKind.Errors)
                        {
                            CollectSpecialCells(area, Excel.XlCellType.xlCellTypeFormulas,
                                Excel.XlSpecialCellsValue.xlErrors, cellKind, functions, matches, ct);
                            CollectSpecialCells(area, Excel.XlCellType.xlCellTypeConstants,
                                Excel.XlSpecialCellsValue.xlErrors, cellKind, functions, matches, ct,
                                formulaErrors: false);
                        }
                        else
                        {
                            var nativeKind = cellKind switch
                            {
                                SpecialCellKind.Formulas => Excel.XlCellType.xlCellTypeFormulas,
                                SpecialCellKind.Constants => Excel.XlCellType.xlCellTypeConstants,
                                SpecialCellKind.Blanks => Excel.XlCellType.xlCellTypeBlanks,
                                SpecialCellKind.Visible => Excel.XlCellType.xlCellTypeVisible,
                                _ => throw new ArgumentOutOfRangeException(nameof(cellKind))
                            };
                            CollectSpecialCells(area, nativeKind, Type.Missing,
                                cellKind, functions, matches, ct);
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref area);
                    }
                }

                foreach (var match in matches.OrderBy(item => item.Row).ThenBy(item => item.Column))
                {
                    result.Areas.Add(match.Address);
                    result.CellCount += match.CellCount;
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref usedRange);
                ComUtilities.Release(ref functions);
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref range);
            }
        });
    }

    private static void CollectBlankCells(
        Excel.Range scope, Excel.Range usedRange, Excel.Worksheet sheet,
        Excel.Application app, List<SpecialCellArea> matches, CancellationToken ct)
    {
        Excel.Range? inside = null;
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        Excel.Range? insideRows = null;
        Excel.Range? insideColumns = null;
        try
        {
            // Native blank discovery omits the part of a request outside UsedRange.
            inside = app.Intersect(scope, usedRange);
            if (inside is null)
            {
                AddSpecialCellArea(scope, matches);
                return;
            }

            if (Convert.ToInt64(inside.CountLarge, CultureInfo.InvariantCulture) == 1)
            {
                if (MatchesCell(inside, SpecialCellKind.Blanks, null))
                {
                    AddSpecialCellArea(inside, matches);
                }
            }
            else
            {
                CollectSpecialCells(inside, Excel.XlCellType.xlCellTypeBlanks,
                    Type.Missing, SpecialCellKind.Blanks, null, matches, ct);
            }

            rows = scope.Rows;
            columns = scope.Columns;
            insideRows = inside.Rows;
            insideColumns = inside.Columns;
            int firstRow = scope.Row;
            int firstColumn = scope.Column;
            int lastRow = firstRow + rows.Count - 1;
            int lastColumn = firstColumn + columns.Count - 1;
            int insideFirstRow = inside.Row;
            int insideFirstColumn = inside.Column;
            int insideLastRow = insideFirstRow + insideRows.Count - 1;
            int insideLastColumn = insideFirstColumn + insideColumns.Count - 1;

            AddBlankRectangle(firstRow, insideFirstRow - 1, firstColumn, lastColumn);
            AddBlankRectangle(insideLastRow + 1, lastRow, firstColumn, lastColumn);
            AddBlankRectangle(insideFirstRow, insideLastRow, firstColumn, insideFirstColumn - 1);
            AddBlankRectangle(insideFirstRow, insideLastRow, insideLastColumn + 1, lastColumn);
        }
        finally
        {
            ComUtilities.Release(ref insideColumns);
            ComUtilities.Release(ref insideRows);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref inside);
        }

        void AddBlankRectangle(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (firstRow > lastRow || firstColumn > lastColumn)
            {
                return;
            }
            ct.ThrowIfCancellationRequested();
            Excel.Range? cells = null;
            Excel.Range? firstCell = null;
            Excel.Range? lastCell = null;
            Excel.Range? rectangle = null;
            try
            {
                cells = sheet.Cells;
                firstCell = cells[firstRow, firstColumn];
                lastCell = cells[lastRow, lastColumn];
                rectangle = sheet.Range[firstCell, lastCell];
                AddSpecialCellArea(rectangle, matches);
            }
            finally
            {
                ComUtilities.Release(ref rectangle);
                ComUtilities.Release(ref lastCell);
                ComUtilities.Release(ref firstCell);
                ComUtilities.Release(ref cells);
            }
        }
    }

    private static void CollectSpecialCells(
        Excel.Range scope, Excel.XlCellType nativeKind, object valueKind,
        SpecialCellKind cellKind, Excel.WorksheetFunction? functions,
        List<SpecialCellArea> matches, CancellationToken ct, bool formulaErrors = true)
    {
        Excel.Range? selected = null;
        Excel.Areas? selectedAreas = null;
        try
        {
            try
            {
                selected = scope.SpecialCells(nativeKind, valueKind);
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
            {
                // This HRESULT is not specific to "no cells". Prove absence before accepting it.
                if (ContainsMatchingCell(scope, cellKind, functions, ct,
                        cellKind == SpecialCellKind.Errors ? formulaErrors : null))
                {
                    throw;
                }
                return;
            }

            selectedAreas = selected.Areas;
            int areaCount = selectedAreas.Count;
            for (int index = 1; index <= areaCount; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Range? area = null;
                try
                {
                    area = selectedAreas[index];
                    AddSpecialCellArea(area, matches);
                }
                finally
                {
                    ComUtilities.Release(ref area);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref selectedAreas);
            ComUtilities.Release(ref selected);
        }
    }

    private static bool ContainsMatchingCell(
        Excel.Range scope, SpecialCellKind cellKind, Excel.WorksheetFunction? functions,
        CancellationToken ct, bool? formulaErrors)
    {
        Excel.Range? cells = null;
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            cells = scope.Cells;
            rows = scope.Rows;
            columns = scope.Columns;
            int rowCount = rows.Count;
            int columnCount = columns.Count;
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? cell = null;
                    try
                    {
                        cell = cells[row, column];
                        if (formulaErrors.HasValue
                            && Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture) != formulaErrors.Value)
                        {
                            continue;
                        }
                        if (MatchesCell(cell, cellKind, functions))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
            }
            return false;
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref cells);
        }
    }

    private static bool MatchesCell(
        Excel.Range cell, SpecialCellKind cellKind, Excel.WorksheetFunction? functions)
    {
        if (cellKind == SpecialCellKind.Visible)
        {
            Excel.Range? row = null;
            Excel.Range? column = null;
            try
            {
                row = cell.EntireRow;
                column = cell.EntireColumn;
                return !Convert.ToBoolean(row.Hidden, CultureInfo.InvariantCulture)
                    && !Convert.ToBoolean(column.Hidden, CultureInfo.InvariantCulture);
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref row);
            }
        }
        if (cellKind == SpecialCellKind.Errors)
        {
            return functions!.IsError(cell);
        }

        bool hasFormula = Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture);
        return cellKind switch
        {
            SpecialCellKind.Formulas => hasFormula,
            SpecialCellKind.Constants => !hasFormula && cell.Value2 is not null,
            SpecialCellKind.Blanks => !hasFormula && cell.Value2 is null,
            _ => throw new ArgumentOutOfRangeException(nameof(cellKind))
        };
    }

    private static void AddSpecialCellArea(Excel.Range area, List<SpecialCellArea> matches) =>
        matches.Add(new SpecialCellArea(area.Address, area.Row, area.Column,
            Convert.ToInt64(area.CountLarge, CultureInfo.InvariantCulture)));

    private sealed record SpecialCellArea(string Address, int Row, int Column, long CellCount);
}

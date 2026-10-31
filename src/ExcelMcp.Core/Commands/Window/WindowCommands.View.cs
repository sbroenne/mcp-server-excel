using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

/// <summary>
/// Worksheet-specific workbook window view operations.
/// </summary>
public partial class WindowCommands
{
    private const int MaxSplitRows = 1_048_575;
    private const int MaxSplitColumns = 16_383;

    /// <inheritdoc />
    public WorksheetViewResult GetView(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                windows = ctx.Book.Windows;
                window = windows[1];
                ActivateWorksheetView(window, sheet);

                return ReadWorksheetView(batch, "get-view", sheetName, window, sheet);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public WorksheetViewResult SetView(IExcelBatch batch, string sheetName, string topLeftCell, string? select = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ValidateViewAddress(topLeftCell, nameof(topLeftCell));
        if (select is not null)
            ValidateViewAddress(select, nameof(select));

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? origin = null;
            Excel.Range? selection = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Panes? panes = null;
            Excel.Pane? pane = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                origin = sheet.Range[topLeftCell];
                ValidateViewRange(ctx.Book, sheet, origin, nameof(topLeftCell), singleCell: true);
                if (select is not null)
                {
                    selection = sheet.Range[select];
                    ValidateViewRange(ctx.Book, sheet, selection, nameof(select), singleCell: false);
                }

                windows = ctx.Book.Windows;
                window = windows[1];
                ct.ThrowIfCancellationRequested();
                ActivateWorksheetView(window, sheet);
                selection?.Select();
                panes = window.Panes;
                pane = panes[panes.Count];
                pane.ScrollRow = origin.Row;
                pane.ScrollColumn = origin.Column;
                return ReadWorksheetView(batch, "set-view", sheetName, window, sheet);
            }
            finally
            {
                ComUtilities.Release(ref pane);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref origin);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidateViewAddress(string address, string parameterName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address, parameterName);
        if (address.IndexOfAny(['!', '[', ']']) >= 0)
            throw new ArgumentException("Use an address on the named worksheet, without a sheet or workbook qualifier.", parameterName);
    }

    private static void ValidateViewRange(
        Excel.Workbook workbook, Excel.Worksheet sheet, Excel.Range range, string parameterName, bool singleCell)
    {
        Excel.Areas? areas = null;
        Excel.Worksheet? rangeSheet = null;
        Excel.Workbook? rangeBook = null;
        try
        {
            areas = range.Areas;
            if (areas.Count != 1)
                throw new ArgumentException("Use one continuous range on the named worksheet.", parameterName);
            if (singleCell && Convert.ToInt64((object)range.CountLarge, System.Globalization.CultureInfo.InvariantCulture) != 1)
                throw new ArgumentException("The top-left position must identify exactly one cell.", parameterName);
            rangeSheet = range.Worksheet;
            rangeBook = (Excel.Workbook)rangeSheet.Parent;
            if (!string.Equals(rangeSheet.Name, sheet.Name, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(rangeBook.FullName, workbook.FullName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The range must belong to the named worksheet in this workbook.", parameterName);
        }
        finally
        {
            ComUtilities.Release(ref rangeBook);
            ComUtilities.Release(ref rangeSheet);
            ComUtilities.Release(ref areas);
        }
    }

    private static WorksheetViewResult ReadWorksheetView(
        IExcelBatch batch, string action, string sheetName, Excel.Window window, Excel.Worksheet sheet)
    {
        Excel.Panes? panes = null;
        Excel.Pane? pane = null;
        Excel.Range? cells = null;
        Excel.Range? origin = null;
        object? selection = null;
        Excel.Range? activeCell = null;
        Excel.Chart? chart = null;
        try
        {
            panes = window.Panes;
            pane = panes[panes.Count];
            cells = sheet.Cells;
            origin = (Excel.Range)cells[pane.ScrollRow, pane.ScrollColumn];
            var result = new WorksheetViewResult
            {
                Action = action,
                FilePath = batch.WorkbookPath,
                SheetName = sheetName,
                TopLeftCell = origin.Address,
                FreezePanes = window.FreezePanes,
                SplitRow = window.SplitRow,
                SplitColumn = window.SplitColumn,
                Zoom = Convert.ToInt32((object)window.Zoom, System.Globalization.CultureInfo.InvariantCulture),
                DisplayGridlines = window.DisplayGridlines,
                DisplayHeadings = window.DisplayHeadings,
                DisplayOutlineSymbols = window.DisplayOutline,
                DisplayFormulas = window.DisplayFormulas
            };
            selection = window.Selection;
            if (selection is Excel.Range range)
            {
                result.SelectionKind = "range";
                result.SelectedRangeAddress = range.Address;
                activeCell = window.ActiveCell;
                result.ActiveCellAddress = activeCell.Address;
            }
            else
            {
                chart = window.ActiveChart;
                result.SelectionKind = chart is not null ? "chart"
                    : selection is Excel.ShapeRange ? "shapes" : "unsupported";
            }
            result.Success = true;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref chart);
            ComUtilities.Release(ref activeCell);
            ComUtilities.Release(ref selection);
            ComUtilities.Release(ref origin);
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref pane);
            ComUtilities.Release(ref panes);
        }
    }

    /// <inheritdoc />
    public OperationResult FreezePanes(IExcelBatch batch, string sheetName, int frozenRows = 0, int frozenColumns = 0)
    {
        ValidatePaneCounts(frozenRows, frozenColumns, requireBoundary: true);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                windows = ctx.Book.Windows;
                window = windows[1];
                ActivateWorksheetView(window, sheet);

                if (window.View == Excel.XlWindowView.xlPageLayoutView)
                {
                    throw new InvalidOperationException(
                        $"Cannot freeze panes on worksheet '{sheetName}' in Page Layout view. " +
                        "Switch the worksheet to Normal view in Excel, then retry. " +
                        "Pane settings and selection were not changed.");
                }

                window.FreezePanes = false;
                if (frozenRows > 0)
                    window.ScrollRow = 1;
                if (frozenColumns > 0)
                    window.ScrollColumn = 1;
                window.SplitRow = frozenRows;
                window.SplitColumn = frozenColumns;
                window.FreezePanes = true;

                var actualFrozen = window.FreezePanes;
                var actualRows = window.SplitRow;
                var actualColumns = window.SplitColumn;
                if (!actualFrozen || actualRows != frozenRows || actualColumns != frozenColumns)
                {
                    throw new InvalidOperationException(
                        $"Excel did not apply the requested frozen panes on worksheet '{sheetName}'. " +
                        $"Requested rows={frozenRows}, columns={frozenColumns}; " +
                        $"actual freezePanes={actualFrozen}, splitRow={actualRows}, splitColumn={actualColumns}. " +
                        "Changes were not rolled back. Inspect the worksheet with window get-view before retrying.");
                }

                return ViewOperationResult(batch, "freeze-panes", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult UnfreezePanes(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                windows = ctx.Book.Windows;
                window = windows[1];
                ActivateWorksheetView(window, sheet);

                window.FreezePanes = false;
                window.SplitRow = 0;
                window.SplitColumn = 0;

                return ViewOperationResult(batch, "unfreeze-panes", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetSplit(IExcelBatch batch, string sheetName, int splitRows = 0, int splitColumns = 0)
    {
        return batch.Execute((ctx, ct) =>
        {
            ValidatePaneCounts(splitRows, splitColumns, requireBoundary: false);

            Excel.Worksheet? sheet = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                windows = ctx.Book.Windows;
                window = windows[1];
                ActivateWorksheetView(window, sheet);

                window.FreezePanes = false;
                window.SplitRow = splitRows;
                window.SplitColumn = splitColumns;

                return ViewOperationResult(batch, "set-split", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetZoom(IExcelBatch batch, string sheetName, int zoom)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (zoom is < 10 or > 400)
            {
                throw new ArgumentOutOfRangeException(nameof(zoom), zoom, "Zoom must be between 10 and 400 percent.");
            }

            Excel.Worksheet? sheet = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                windows = ctx.Book.Windows;
                window = windows[1];
                ActivateWorksheetView(window, sheet);
                window.Zoom = zoom;

                return ViewOperationResult(batch, "set-zoom", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetDisplayOptions(
        IExcelBatch batch,
        string sheetName,
        bool? showGridlines = null,
        bool? showHeadings = null,
        bool? showOutlineSymbols = null,
        bool? showFormulas = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (!showGridlines.HasValue
                && !showHeadings.HasValue
                && !showOutlineSymbols.HasValue
                && !showFormulas.HasValue)
            {
                throw new ArgumentException("Provide at least one display option to change.");
            }

            Excel.Worksheet? sheet = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                windows = ctx.Book.Windows;
                window = windows[1];
                ActivateWorksheetView(window, sheet);

                if (showGridlines.HasValue)
                {
                    window.DisplayGridlines = showGridlines.Value;
                }

                if (showHeadings.HasValue)
                {
                    window.DisplayHeadings = showHeadings.Value;
                }

                if (showOutlineSymbols.HasValue)
                {
                    window.DisplayOutline = showOutlineSymbols.Value;
                }

                if (showFormulas.HasValue)
                {
                    window.DisplayFormulas = showFormulas.Value;
                }

                return ViewOperationResult(batch, "set-display-options", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static Excel.Worksheet FindRequiredSheet(Excel.Workbook workbook, string sheetName)
    {
        return ComUtilities.FindSheet(workbook, sheetName)
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
    }

    private static void ActivateWorksheetView(Excel.Window window, Excel.Worksheet sheet)
    {
        window.Activate();
        sheet.Activate();
    }

    private static void ValidatePaneCounts(int rows, int columns, bool requireBoundary)
    {
        if (rows < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), rows, "Row count cannot be negative.");
        }

        if (rows > MaxSplitRows)
        {
            throw new ArgumentOutOfRangeException(
                nameof(rows),
                rows,
                $"Row count cannot exceed {MaxSplitRows}.");
        }

        if (columns < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columns), columns, "Column count cannot be negative.");
        }

        if (columns > MaxSplitColumns)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columns),
                columns,
                $"Column count cannot exceed {MaxSplitColumns}.");
        }

        if (requireBoundary && rows == 0 && columns == 0)
        {
            throw new ArgumentException("At least one frozen row or frozen column is required.");
        }
    }

    private static OperationResult ViewOperationResult(IExcelBatch batch, string action, string sheetName)
    {
        return new OperationResult
        {
            Success = true,
            Action = action,
            FilePath = batch.WorkbookPath,
            Message = $"Worksheet view updated for '{sheetName}'."
        };
    }
}

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
        var result = new WorksheetViewResult
        {
            Action = "get-view",
            FilePath = batch.WorkbookPath,
            SheetName = sheetName
        };

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

                result.FreezePanes = window.FreezePanes;
                result.SplitRow = window.SplitRow;
                result.SplitColumn = window.SplitColumn;
                result.Zoom = Convert.ToInt32(window.Zoom, System.Globalization.CultureInfo.InvariantCulture);
                result.DisplayGridlines = window.DisplayGridlines;
                result.DisplayHeadings = window.DisplayHeadings;
                result.DisplayOutlineSymbols = window.DisplayOutline;
                result.DisplayFormulas = window.DisplayFormulas;
                result.Success = true;
                return result;
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
    public OperationResult FreezePanes(IExcelBatch batch, string sheetName, int frozenRows = 0, int frozenColumns = 0)
    {
        return batch.Execute((ctx, ct) =>
        {
            ValidatePaneCounts(frozenRows, frozenColumns, requireBoundary: true);

            Excel.Worksheet? sheet = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Range? boundaryCell = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                windows = ctx.Book.Windows;
                window = windows[1];
                ActivateWorksheetView(window, sheet);

                window.FreezePanes = false;
                boundaryCell = sheet.Cells[frozenRows + 1, frozenColumns + 1];
                boundaryCell.Select();
                window.FreezePanes = true;

                return ViewOperationResult(batch, "freeze-panes", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref boundaryCell);
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

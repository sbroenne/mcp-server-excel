using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

public partial class WindowCommands
{
    /// <inheritdoc />
    public WindowContextResult GetContext(IExcelBatch batch)
    {
        return batch.Execute((context, ct) =>
        {
            Excel.Windows? windows = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                windows = context.Book.Windows;
                var result = new WindowContextResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "get-context",
                    IsApplicationVisible = context.App.Visible,
                    Availability = windows.Count == 0 ? "no-windows" : "available"
                };
                for (int index = 1; index <= windows.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Window? window = null;
                    object? sheet = null;
                    object? selection = null;
                    Excel.Range? activeCell = null;
                    Excel.Chart? chart = null;
                    try
                    {
                        window = windows[index];
                        var item = new WorkbookWindowContext
                        {
                            WindowNumber = window.WindowNumber,
                            IsVisible = window.Visible
                        };
                        sheet = window.ActiveSheet;
                        if (sheet is Excel.Worksheet worksheet)
                        {
                            item.SheetKind = "worksheet";
                            item.SheetName = worksheet.Name;
                        }
                        else if (sheet is Excel.Chart chartSheet)
                        {
                            item.SheetKind = "chart-sheet";
                            item.SheetName = chartSheet.Name;
                        }
                        else
                        {
                            item.SheetKind = "unsupported";
                        }
                        try
                        {
                            selection = window.Selection;
                            chart = window.ActiveChart;
                            item.ActiveChartName = chart?.Name;
                            if (selection is Excel.Range range)
                            {
                                item.SelectionKind = "range";
                                item.SelectedRangeAddress = range.Address;
                                activeCell = window.ActiveCell;
                                item.ActiveCellAddress = activeCell?.Address;
                            }
                            else if (chart is not null)
                            {
                                item.SelectionKind = "chart";
                            }
                            else if (selection is Excel.ShapeRange)
                            {
                                item.SelectionKind = "shapes";
                            }
                            else
                            {
                                item.SelectionKind = "unsupported";
                                item.SelectionDiagnostic = "The native selection does not expose a supported range/chart/shape interface.";
                            }
                        }
                        catch (COMException exception) when (exception.HResult == unchecked((int)0x800A03EC))
                        {
                            item.SelectionKind = "unavailable";
                            item.SelectedRangeAddress = null;
                            item.ActiveCellAddress = null;
                            item.ActiveChartName = null;
                            item.SelectionDiagnostic = $"Native window selection is unavailable: {exception.Message}";
                        }
                        result.Windows.Add(item);
                    }
                    finally
                    {
                        ComUtilities.Release(ref activeCell);
                        ComUtilities.Release(ref chart);
                        ComUtilities.Release(ref selection);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref window);
                    }
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref windows);
            }
        });
    }
}

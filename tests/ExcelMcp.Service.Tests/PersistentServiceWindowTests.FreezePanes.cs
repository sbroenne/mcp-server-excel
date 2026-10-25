using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceWindowTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FreezePanes_ReplacementWorksheetWithAnotherSheetActive_FreezesRequestedRow(bool show)
    {
        var batch = _fixture.BatchToken;
        var sheets = _fixture.CreateCommands<ISheetCommands>();
        var ranges = _fixture.CreateCommands<IRangeCommands>();
        var formats = _fixture.CreateCommands<IRangeFormatCommands>();
        var original = _fixture.CreateTestSheet(batch);
        var replacement = _fixture.CreateTestSheet(batch);
        var other = _fixture.CreateTestSheet(batch);

        if (show)
            RequireSuccess(_commands.Show(batch));

        RequireSuccess(sheets.Move(batch, replacement, beforeSheet: original));
        RequireSuccess(ranges.SetValues(batch, replacement, "A1:B3",
            [["Code", "Label"], [1, "First"], [2, "Second"]]));
        RequireSuccess(ranges.SetFormulas(batch, other, "A1", [[$"='{replacement}'!B2"]]));
        RequireSuccess(sheets.Delete(batch, original));
        _fixture.ForgetSheet(original);
        RequireSuccess(sheets.Rename(batch, replacement, original));
        _fixture.RenameTrackedSheet(replacement, original);
        RequireSuccess(formats.Format(batch, original, ["A1:B1"], new CellFormatOptions { Bold = true }));
        RequireSuccess(formats.AutoFitColumns(batch, original, "A:B"));
        RequireSuccess(formats.SetColumnWidth(batch, original, "B:B", 24));
        RequireSuccess(formats.AutoFitRows(batch, original, "1:3"));

        RequireSuccess(_commands.FreezePanes(batch, other, frozenRows: 3, frozenColumns: 2));
        RequireSuccess(_commands.SetZoom(batch, other, 85));
        var otherBefore = RequireSuccess(_commands.GetView(batch, other));
        Assert.Equal(other, Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows).SheetName);

        RequireSuccess(_commands.FreezePanes(batch, original, frozenRows: 1, frozenColumns: 0));

        Assert.Equal(original, Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows).SheetName);
        Assert.Equal((true, 1, 0), ReadActivePaneState());
        var otherAfter = RequireSuccess(_commands.GetView(batch, other));
        Assert.Equal(otherBefore.FreezePanes, otherAfter.FreezePanes);
        Assert.Equal(otherBefore.SplitRow, otherAfter.SplitRow);
        Assert.Equal(otherBefore.SplitColumn, otherAfter.SplitColumn);
        Assert.Equal(otherBefore.Zoom, otherAfter.Zoom);
    }

    [Theory]
    [InlineData(false, 1, 0)]
    [InlineData(false, 0, 2)]
    [InlineData(false, 2, 1)]
    [InlineData(true, 1, 0)]
    [InlineData(true, 0, 2)]
    [InlineData(true, 2, 1)]
    public void FreezePanes_ExistingPaneBoundary_ReplacesRequestedCounts(bool frozen, int rows, int columns)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        if (frozen)
            RequireSuccess(_commands.FreezePanes(batch, sheet, frozenRows: 4, frozenColumns: 3));
        else
            RequireSuccess(_commands.SetSplit(batch, sheet, splitRows: 4, splitColumns: 3));
        Assert.Equal((frozen, 4, 3), ReadActivePaneState());

        RequireSuccess(_commands.FreezePanes(batch, sheet, frozenRows: rows, frozenColumns: columns));

        Assert.Equal((true, rows, columns), ReadActivePaneState());
    }

    [Theory]
    [InlineData(Excel.XlWindowView.xlNormalView)]
    [InlineData(Excel.XlWindowView.xlPageBreakPreview)]
    public void FreezePanes_ScrolledWorksheet_FreezesFromTopLeftWithoutChangingSelection(Excel.XlWindowView view)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Worksheet? worksheet = null;
            Excel.Range? selection = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                worksheet = ComUtilities.FindSheet(context.Book, sheet);
                Assert.NotNull(worksheet);
                window.View = view;
                selection = worksheet.Range["H30"];
                selection.Select();
                window.ScrollRow = 20;
                window.ScrollColumn = 5;
                Assert.Equal(20, window.ScrollRow);
                Assert.Equal(5, window.ScrollColumn);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });

        RequireSuccess(_commands.FreezePanes(batch, sheet, frozenRows: 1, frozenColumns: 2));

        Assert.Equal((true, 1, 2), ReadActivePaneState());
        Assert.Equal("$H$30", Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows).SelectedRangeAddress);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Panes? panes = null;
            Excel.Pane? topLeft = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                panes = window.Panes;
                topLeft = panes[1];
                Assert.Equal(view, window.View);
                Assert.Equal(1, topLeft.ScrollRow);
                Assert.Equal(1, topLeft.ScrollColumn);
            }
            finally
            {
                ComUtilities.Release(ref topLeft);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
    }

    [Fact]
    public async Task FreezePanes_PageLayoutView_ExplainsFailureWithoutChangingSelection()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Worksheet? worksheet = null;
            Excel.Range? selection = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                worksheet = ComUtilities.FindSheet(context.Book, sheet);
                Assert.NotNull(worksheet);
                window.View = Excel.XlWindowView.xlPageLayoutView;
                selection = worksheet.Range["D7"];
                selection.Select();
                Assert.Equal(Excel.XlWindowView.xlPageLayoutView, window.View);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
        var before = Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows);
        Assert.Equal("$D$7", before.SelectedRangeAddress);
        Assert.Equal((false, 0, 0), ReadActivePaneState());

        var failure = await _fixture.SendForFailureAsync("window.freeze-panes",
            new { sheetName = sheet, frozenRows = 1, frozenColumns = 0 });

        Assert.Equal("window.freeze-panes", failure.Command);
        Assert.Equal(nameof(InvalidOperationException), failure.ExceptionType);
        Assert.Contains("Page Layout", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Normal", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Pane settings and selection were not changed", failure.ErrorMessage, StringComparison.Ordinal);
        var after = Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows);
        Assert.Equal(before.SheetName, after.SheetName);
        Assert.Equal(before.SelectedRangeAddress, after.SelectedRangeAddress);
        Assert.Equal((false, 0, 0), ReadActivePaneState());
        Assert.Equal(Excel.XlWindowView.xlPageLayoutView, _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                return window.View;
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        }));
    }

    private (bool Frozen, int Rows, int Columns) ReadActivePaneState() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                return (window.FreezePanes, window.SplitRow, window.SplitColumn);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
}

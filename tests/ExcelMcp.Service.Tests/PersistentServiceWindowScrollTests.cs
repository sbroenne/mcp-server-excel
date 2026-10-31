using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Window")]
public sealed class PersistentServiceWindowScrollTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWindowTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void GetView_ReportsActualScrollSelectionAndActiveCell()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.GetView(_fixture.BatchToken, sheet));
        SetNativeView(sheet, "H30:I32", 20, 5);

        var view = RequireSuccess(_commands.GetView(_fixture.BatchToken, sheet));
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(view, ServiceProtocol.JsonOptions));

        Assert.Equal("$E$20", json.RootElement.GetProperty("topLeftCell").GetString());
        Assert.Equal("range", json.RootElement.GetProperty("selectionKind").GetString());
        Assert.Equal("$H$30:$I$32", json.RootElement.GetProperty("selectedRangeAddress").GetString());
        Assert.Equal("$H$30", json.RootElement.GetProperty("activeCellAddress").GetString());
        AssertNativeView(20, 5, "$H$30:$I$32", "$H$30");
    }

    [Theory]
    [InlineData(false, "A1", "A1", 1, 1, "$A$1", "$A$1")]
    [InlineData(true, "E20", "H30:I32", 20, 5, "$H$30:$I$32", "$H$30")]
    [InlineData(false, "C7", null, 7, 3, "$H$30:$I$32", "$H$30")]
    public void SetView_ChangesTargetAndPreservesOtherSheet(
        bool show, string origin, string? select, int row, int column, string address, string activeAddress)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var other = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        SetNativeView(sheet, "H30:I32", 31, 4);
        RequireSuccess(_commands.SetZoom(batch, sheet, 135));
        RequireSuccess(_commands.SetDisplayOptions(batch, sheet, showGridlines: false));
        RequireSuccess(_commands.GetView(batch, other));
        SetNativeView(other, "L40", 25, 6);
        if (show)
            RequireSuccess(_commands.Show(batch));
        var visibility = RequireSuccess(_commands.GetInfo(batch)).IsVisible;

        var result = RequireSuccess(_commands.SetView(batch, sheet, origin, select));

        Assert.Equal("set-view", result.Action);
        Assert.Equal(sheet, result.SheetName);
        Assert.Equal($"${(char)('A' + column - 1)}${row}", result.TopLeftCell);
        Assert.Equal("range", result.SelectionKind);
        Assert.Equal(address, result.SelectedRangeAddress);
        Assert.Equal(activeAddress, result.ActiveCellAddress);
        Assert.Equal(135, result.Zoom);
        Assert.False(result.DisplayGridlines);
        Assert.Equal(visibility, RequireSuccess(_commands.GetInfo(batch)).IsVisible);
        Assert.Equal(sheet, Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows).SheetName);
        AssertNativeView(row, column, address, activeAddress);
        var read = RequireSuccess(_commands.GetView(batch, sheet));
        Assert.Equal(result.TopLeftCell, read.TopLeftCell);
        Assert.Equal(result.SelectedRangeAddress, read.SelectedRangeAddress);
        RequireSuccess(_commands.GetView(batch, other));
        AssertNativeView(25, 6, "$L$40", "$L$40");
    }

    [Theory]
    [InlineData(true, 2, 0)]
    [InlineData(true, 0, 2)]
    [InlineData(true, 2, 2)]
    [InlineData(false, 2, 0)]
    [InlineData(false, 0, 2)]
    [InlineData(false, 2, 2)]
    public void SetView_PreservesPaneBoundariesAndControlsLastPane(bool frozen, int rows, int columns)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        if (frozen)
            RequireSuccess(_commands.FreezePanes(batch, sheet, rows, columns));
        else
            RequireSuccess(_commands.SetSplit(batch, sheet, rows, columns));
        SetNativeView(sheet, "H30", 20, 5);
        var firstBefore = ReadFirstPane();

        var result = RequireSuccess(_commands.SetView(batch, sheet, "G25"));

        Assert.Equal("$G$25", result.TopLeftCell);
        Assert.Equal(frozen, result.FreezePanes);
        Assert.Equal(rows, result.SplitRow);
        Assert.Equal(columns, result.SplitColumn);
        Assert.Equal("$H$30", result.SelectedRangeAddress);
        AssertNativeView(25, 7, "$H$30", "$H$30");
        Assert.Equal((rows > 0 ? firstBefore.Row : 25, columns > 0 ? firstBefore.Column : 7), ReadFirstPane());
    }

    [Theory]
    [InlineData(2, 0, "$A$3")]
    [InlineData(0, 2, "$C$1")]
    [InlineData(2, 2, "$C$3")]
    public void SetView_FrozenPane_AdjustsOriginWithoutRemovingFrozenCells(int rows, int columns, string expected)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.FreezePanes(batch, sheet, rows, columns));
        SetNativeView(sheet, "H30", 20, 5);

        var result = RequireSuccess(_commands.SetView(batch, sheet, "A1", "A1"));

        Assert.Equal(expected, result.TopLeftCell);
        Assert.True(result.FreezePanes);
        Assert.Equal(rows, result.SplitRow);
        Assert.Equal(columns, result.SplitColumn);
        Assert.Equal("$A$1", result.SelectedRangeAddress);
        Assert.Equal("$A$1", result.ActiveCellAddress);
    }

    [Theory]
    [InlineData("", null)]
    [InlineData("A1:B2", null)]
    [InlineData("A1,B3", null)]
    [InlineData("Other!A1", null)]
    [InlineData("[Other.xlsx]Sheet1!A1", null)]
    [InlineData("XFE1", null)]
    [InlineData("A1048577", null)]
    [InlineData("not_a_cell", null)]
    [InlineData("A1", "")]
    [InlineData("A1", "A1,B3")]
    [InlineData("A1", "Other!A1")]
    public async Task SetView_InvalidAddress_PreservesView(string origin, string? select)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        SetNativeView(sheet, "H30", 20, 5);

        var failure = await _fixture.SendForFailureAsync("window.set-view",
            new { sheetName = sheet, topLeftCell = origin, select });

        Assert.False(failure.Success);
        Assert.False(string.IsNullOrWhiteSpace(failure.ErrorMessage));
        AssertNativeView(20, 5, "$H$30", "$H$30");
    }

    [Fact]
    public async Task SetView_MissingSheet_PreservesActiveView()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        SetNativeView(sheet, "H30", 20, 5);

        var failure = await _fixture.SendForFailureAsync("window.set-view",
            new { sheetName = "MissingSheet", topLeftCell = "A1", select = "A1" });

        Assert.Contains("not found", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(sheet, Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows).SheetName);
        AssertNativeView(20, 5, "$H$30", "$H$30");
    }

    [Fact]
    public void SetView_HiddenRowsAndColumns_ReportsNativePosition()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        SetNativeView(sheet, "H30", 20, 5);
        ModifyNativeSheet(sheet, worksheet =>
        {
            Excel.Range? rowCells = null;
            Excel.Range? columnCells = null;
            Excel.Range? rows = null;
            Excel.Range? columns = null;
            try
            {
                rowCells = worksheet.Range["A5:A6"];
                columnCells = worksheet.Range["C1:D1"];
                rows = rowCells.EntireRow;
                columns = columnCells.EntireColumn;
                rows.Hidden = true;
                columns.Hidden = true;
            }
            finally
            {
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref columnCells);
                ComUtilities.Release(ref rowCells);
            }
        });

        var result = RequireSuccess(_commands.SetView(batch, sheet, "C5"));

        Assert.Equal("$B$4", result.TopLeftCell);
        AssertNativeView(4, 2, "$H$30", "$H$30");
    }

    [Fact]
    public void SetView_MergedCell_ReportsActualSelection()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        SetNativeView(sheet, "H30", 20, 5);
        ModifyNativeSheet(sheet, worksheet =>
        {
            Excel.Range? merged = null;
            try
            {
                merged = worksheet.Range["A1:B2"];
                merged.Merge();
            }
            finally
            {
                ComUtilities.Release(ref merged);
            }
        });

        var result = RequireSuccess(_commands.SetView(batch, sheet, "A1", "A1"));

        Assert.Equal("$A$1", result.TopLeftCell);
        Assert.Equal("$A$1:$B$2", result.SelectedRangeAddress);
        Assert.Equal("$A$1", result.ActiveCellAddress);
        AssertNativeView(1, 1, "$A$1:$B$2", "$A$1");
    }

    [Fact]
    public void SetView_OmittedSelection_PreservesSelectedShapeWithoutStaleRange()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        SetNativeView(sheet, "H30", 20, 5);
        ModifyNativeSheet(sheet, worksheet =>
        {
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                shapes = worksheet.Shapes;
                shape = shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 20, 30, 80, 40);
                shape.Select();
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
            }
        });
        var before = RequireSuccess(_commands.GetView(batch, sheet));
        Assert.Equal("unsupported", before.SelectionKind);
        Assert.Null(before.SelectedRangeAddress);
        Assert.Null(before.ActiveCellAddress);

        var result = RequireSuccess(_commands.SetView(batch, sheet, "E20"));

        Assert.Equal("$E$20", result.TopLeftCell);
        Assert.Equal(before.SelectionKind, result.SelectionKind);
        Assert.Null(result.SelectedRangeAddress);
        Assert.Null(result.ActiveCellAddress);
        Assert.Equal(before.SelectionKind, Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows).SelectionKind);

        foreach (var response in new[]
        {
            _fixture.Send("window.get-view", new { sheetName = sheet }),
            _fixture.Send("window.set-view", new { sheetName = sheet, topLeftCell = "E20" })
        })
        {
            Assert.True(response.Success, response.ErrorMessage);
            using var json = JsonDocument.Parse(response.Result!);
            Assert.Equal("unsupported", json.RootElement.GetProperty("selectionKind").GetString());
            Assert.Equal("$E$20", json.RootElement.GetProperty("topLeftCell").GetString());
            Assert.Equal(JsonValueKind.Null, json.RootElement.GetProperty("selectedRangeAddress").ValueKind);
            Assert.Equal(JsonValueKind.Null, json.RootElement.GetProperty("activeCellAddress").ValueKind);
        }
    }

    [Fact]
    public async Task SetView_HiddenSheet_ReportsFailureAndRetainsOriginalSheetView()
    {
        var batch = _fixture.BatchToken;
        var hidden = _fixture.CreateTestSheet(batch);
        var active = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, active));
        SetNativeView(active, "H30", 20, 5);
        ModifyNativeSheet(hidden, worksheet => worksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden);
        RequireSuccess(_commands.GetView(batch, active));
        SetNativeView(active, "H30", 20, 5);

        var failure = await _fixture.SendForFailureAsync("window.set-view",
            new { sheetName = hidden, topLeftCell = "A1", select = "A1" });

        Assert.False(failure.Success);
        Assert.Equal("Sheet1", Assert.Single(RequireSuccess(_commands.GetContext(batch)).Windows).SheetName);
        RequireSuccess(_commands.GetView(batch, active));
        AssertNativeView(20, 5, "$H$30", "$H$30");
    }

    [Fact]
    public void SetView_ProtectedWorksheet_UsesNativeSelectionBehavior()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.GetView(batch, sheet));
        SetNativeView(sheet, "H30", 20, 5);
        ModifyNativeSheet(sheet, worksheet =>
        {
            worksheet.Protect();
            worksheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;
        });
        PersistentServiceCleanupFailures.Run(() =>
        {
            var result = RequireSuccess(_commands.SetView(batch, sheet, "A1", "A1"));
            Assert.Equal("$A$1", result.TopLeftCell);
            Assert.Equal("$A$1", result.SelectedRangeAddress);
            Assert.Equal("$A$1", result.ActiveCellAddress);
            AssertNativeView(1, 1, "$A$1", "$A$1");
            ModifyNativeSheet(sheet, worksheet =>
            {
                Assert.True(worksheet.ProtectContents);
                Assert.Equal(Excel.XlEnableSelection.xlNoSelection, worksheet.EnableSelection);
            });
        }, () => ModifyNativeSheet(sheet, worksheet => worksheet.Unprotect()));
    }

    private void ModifyNativeSheet(string name, Action<Excel.Worksheet> change)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, name);
                Assert.NotNull(sheet);
                change(sheet);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private (int Row, int Column) ReadFirstPane()
    {
        return _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Panes? panes = null;
            Excel.Pane? pane = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                panes = window.Panes;
                pane = panes[1];
                return (pane.ScrollRow, pane.ScrollColumn);
            }
            finally
            {
                ComUtilities.Release(ref pane);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
    }

    private void SetNativeView(string sheetName, string address, int row, int column)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            Excel.Panes? panes = null;
            Excel.Pane? pane = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                window.Activate();
                sheet.Activate();
                selection = sheet.Range[address];
                selection.Select();
                panes = window.Panes;
                pane = panes[panes.Count];
                pane.ScrollRow = row;
                pane.ScrollColumn = column;
                Assert.Equal(row, pane.ScrollRow);
                Assert.Equal(column, pane.ScrollColumn);
            }
            finally
            {
                ComUtilities.Release(ref pane);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
    }

    private void AssertNativeView(int row, int column, string address, string activeAddress)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            Excel.Panes? panes = null;
            Excel.Pane? pane = null;
            object? selection = null;
            Excel.Range? activeCell = null;
            try
            {
                windows = context.Book.Windows;
                window = windows[1];
                panes = window.Panes;
                pane = panes[panes.Count];
                Assert.Equal(row, pane.ScrollRow);
                Assert.Equal(column, pane.ScrollColumn);
                selection = window.Selection;
                Assert.True(selection is Excel.Range);
                Assert.Equal(address, ((Excel.Range)selection).Address);
                activeCell = window.ActiveCell;
                Assert.Equal(activeAddress, activeCell.Address);
            }
            finally
            {
                ComUtilities.Release(ref activeCell);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref pane);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
    }
}

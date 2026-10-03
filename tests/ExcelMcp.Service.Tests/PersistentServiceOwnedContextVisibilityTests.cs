using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Window")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceOwnedContextVisibilityTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Context_ReadsOwnedHiddenWindowWithoutChangingSelection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["C5"];
                sheet.Activate();
                cell.Select();
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
        var response = _fixture.Send("window.get-context", new { });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal("available", result.RootElement.GetProperty("availability").GetString());
        var window = Assert.Single(result.RootElement.GetProperty("windows").EnumerateArray());
        Assert.Equal(sheetName, window.GetProperty("sheetName").GetString());
        Assert.Equal("range", window.GetProperty("selectionKind").GetString());
        Assert.Equal("$C$5", window.GetProperty("selectedRangeAddress").GetString());
        Assert.Equal("$C$5", window.GetProperty("activeCellAddress").GetString());
    }

    [Theory]
    [InlineData("rows", "A2:C3", "A2:A3", 25d)]
    [InlineData("columns", "B1:C3", "B1:C1", 18d)]
    public void Visibility_HideShowPreservesNativeDimensions(string axis, string rangeAddress,
        string inspect, double size)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        if (axis == "rows")
            _fixture.Send("rangeformat.set-row-height", new { sheetName, rangeAddress, rowHeight = size });
        else
            _fixture.Send("rangeformat.set-column-width", new { sheetName, rangeAddress, columnWidth = size });
        _fixture.Send("rangeformat.set-visibility", new { sheetName, rangeAddress, axis, hidden = true });
        var hidden = _fixture.Send("rangeformat.get-visibility", new { sheetName, rangeAddress = inspect, axis });
        using (var read = JsonDocument.Parse(hidden.Result!))
        {
            Assert.Equal(2, read.RootElement.GetProperty("items").GetArrayLength());
            Assert.All(read.RootElement.GetProperty("items").EnumerateArray(), item =>
            {
                Assert.True(item.GetProperty("hidden").GetBoolean());
                Assert.Equal("undetermined", item.GetProperty("hiddenCause").GetString());
            });
        }
        _fixture.Send("rangeformat.set-visibility", new { sheetName, rangeAddress, axis, hidden = false });
        var shown = _fixture.Send("rangeformat.get-visibility", new { sheetName, rangeAddress = inspect, axis });
        using var result = JsonDocument.Parse(shown.Result!);
        Assert.All(result.RootElement.GetProperty("items").EnumerateArray(), item =>
        {
            Assert.False(item.GetProperty("hidden").GetBoolean());
            Assert.Equal(size, item.GetProperty("size").GetDouble(), precision: 1);
        });
    }

    [Fact]
    public void Visibility_ExactDisjointScopeDoesNotHideGapsOrDoubleCount()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("rangeformat.set-visibility", new
        {
            sheetName,
            rangeAddress = "A2,A4:A5,A5",
            axis = "rows",
            hidden = true
        });
        var response = _fixture.Send("rangeformat.get-visibility", new
        {
            sheetName,
            rangeAddress = "A2:A5",
            axis = "rows"
        });
        using var result = JsonDocument.Parse(response.Result!);
        var items = result.RootElement.GetProperty("items");
        Assert.Equal(4, items.GetArrayLength());
        Assert.True(items[0].GetProperty("hidden").GetBoolean());
        Assert.False(items[1].GetProperty("hidden").GetBoolean());
        Assert.True(items[2].GetProperty("hidden").GetBoolean());
        Assert.True(items[3].GetProperty("hidden").GetBoolean());
        var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A3");
        Assert.True(values.Success);
        Assert.Null(values.Values[0][0]);
    }

    [Fact]
    public void NativeOwnedWindowProbe_ReadsSelectionFromWorkbookWindow()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            object? selected = null;
            Excel.Range? activeCell = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["C5"];
                sheet.Activate();
                cell.Select();
                windows = context.Book.Windows;
                window = windows[1];
                selected = window.Selection;
                activeCell = window.ActiveCell;
                var range = selected as Excel.Range;
                Assert.NotNull(range);
                Assert.Equal("$C$5", range.Address);
                Assert.Equal("$C$5", activeCell.Address);
                Assert.True(Convert.ToInt32(window.WindowNumber, CultureInfo.InvariantCulture) > 0);
            }
            finally
            {
                ComUtilities.Release(ref activeCell);
                ComUtilities.Release(ref selected);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    [Fact]
    public void Context_DoesNotReportAnotherWorkbooksApplicationSelection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Excel.Workbook? foreign = null;
        try
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? ownSheet = null;
                Excel.Range? ownSelection = null;
                Excel.Workbooks? books = null;
                Excel.Sheets? foreignSheets = null;
                Excel.Worksheet? foreignSheet = null;
                Excel.Range? foreignSelection = null;
                try
                {
                    ownSheet = ComUtilities.FindSheet(context.Book, sheetName);
                    Assert.NotNull(ownSheet);
                    ownSelection = ownSheet.Range["C5"];
                    ownSheet.Activate();
                    ownSelection.Select();
                    books = context.App.Workbooks;
                    foreign = books.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    foreignSheets = foreign.Worksheets;
                    foreignSheet = (Excel.Worksheet)foreignSheets[1];
                    foreignSelection = foreignSheet.Range["D8"];
                    foreignSheet.Activate();
                    foreignSelection.Select();
                }
                finally
                {
                    ComUtilities.Release(ref foreignSelection);
                    ComUtilities.Release(ref foreignSheet);
                    ComUtilities.Release(ref foreignSheets);
                    ComUtilities.Release(ref books);
                    ComUtilities.Release(ref ownSelection);
                    ComUtilities.Release(ref ownSheet);
                }
            });
            var response = _fixture.Send("window.get-context", new { });
            using var result = JsonDocument.Parse(response.Result!);
            var window = Assert.Single(result.RootElement.GetProperty("windows").EnumerateArray());
            Assert.Equal(sheetName, window.GetProperty("sheetName").GetString());
            Assert.Equal("$C$5", window.GetProperty("selectedRangeAddress").GetString());
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Range? selected = null;
                try
                {
                    selected = context.App.Selection as Excel.Range;
                    Assert.NotNull(selected);
                    Assert.Equal("$D$8", selected.Address);
                }
                finally
                {
                    ComUtilities.Release(ref selected);
                }
            });
        }
        finally
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                try
                {
                    foreign?.Close(SaveChanges: false);
                }
                finally
                {
                    ComUtilities.Release(ref foreign);
                    context.Book.Activate();
                }
            });
        }
    }

    [Fact]
    public void Visibility_ReportsFilterContextWithoutInventingHiddenCause()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A3",
            [["Flag"], ["keep"], ["drop"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            object? outcome = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range["A1:A3"];
                outcome = range.AutoFilter(Field: 1, Criteria1: "keep");
            }
            finally
            {
                ComUtilities.Release(ref outcome);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
        var response = _fixture.Send("rangeformat.get-visibility", new
        {
            sheetName,
            rangeAddress = "A1:A3",
            axis = "rows"
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("sheetFilterMode").GetBoolean());
        var rows = result.RootElement.GetProperty("items");
        Assert.False(rows[0].GetProperty("withinWorksheetAutoFilterDataRows").GetBoolean());
        Assert.True(rows[2].GetProperty("withinWorksheetAutoFilterDataRows").GetBoolean());
        Assert.True(rows[2].GetProperty("hidden").GetBoolean());
        Assert.Equal("undetermined", rows[2].GetProperty("hiddenCause").GetString());
    }

    [Fact]
    public void Visibility_NamedColumnScopeReportsEachDimensionOnce()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Visibility_{Guid.NewGuid():N}";
        _fixture.Send("namedrange.create", new { name, reference = $"'{sheetName}'!$B$2:$D$4" });
        _fixture.RegisterNamedRangeForCleanup(name);
        _fixture.Send("rangeformat.set-visibility", new
        {
            sheetName = string.Empty,
            rangeAddress = name,
            axis = "columns",
            hidden = true
        });
        var response = _fixture.Send("rangeformat.get-visibility", new
        {
            sheetName = string.Empty,
            rangeAddress = name,
            axis = "columns"
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(sheetName, result.RootElement.GetProperty("sheetName").GetString());
        var items = result.RootElement.GetProperty("items");
        Assert.Equal(3, items.GetArrayLength());
        Assert.Equal([2, 3, 4], items.EnumerateArray().Select(item => item.GetProperty("index").GetInt32()));
        Assert.All(items.EnumerateArray(), item =>
        {
            Assert.True(item.GetProperty("hidden").GetBoolean());
            Assert.Equal("character-width", item.GetProperty("sizeUnit").GetString());
        });
        var gaps = _fixture.Send("rangeformat.get-visibility", new
        {
            sheetName,
            rangeAddress = "A1,E1",
            axis = "columns"
        });
        using var gapResult = JsonDocument.Parse(gaps.Result!);
        Assert.All(gapResult.RootElement.GetProperty("items").EnumerateArray(),
            item => Assert.False(item.GetProperty("hidden").GetBoolean()));
    }

    [Fact]
    public async Task Visibility_ProtectionFailureLeavesDimensionsUnchanged()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                sheet.Protect();
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
        try
        {
            var failure = await _fixture.SendForFailureAsync("rangeformat.set-visibility", new
            {
                sheetName,
                rangeAddress = "A2:A4",
                axis = "rows",
                hidden = true
            });
            Assert.False(failure.Success);
            Assert.False(string.IsNullOrEmpty(failure.ErrorMessage));
            var response = _fixture.Send("rangeformat.get-visibility", new
            {
                sheetName,
                rangeAddress = "A1:A5",
                axis = "rows"
            });
            using var result = JsonDocument.Parse(response.Result!);
            Assert.All(result.RootElement.GetProperty("items").EnumerateArray(),
                item => Assert.False(item.GetProperty("hidden").GetBoolean()));
        }
        finally
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? sheet = null;
                try
                {
                    sheet = ComUtilities.FindSheet(context.Book, sheetName);
                    Assert.NotNull(sheet);
                    sheet.Unprotect();
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
    }

    [Fact]
    public void Context_ActiveChartDoesNotFabricateARangeSelection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ChartObjects? charts = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.ChartArea? area = null;
            object? selectionOutcome = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                sheet.Activate();
                charts = (Excel.ChartObjects)sheet.ChartObjects();
                chartObject = charts.Add(10, 10, 200, 100);
                chartObject.Name = "ContextChart";
                chartObject.Activate();
                chart = chartObject.Chart;
                area = chart.ChartArea;
                selectionOutcome = area.Select();
            }
            finally
            {
                ComUtilities.Release(ref selectionOutcome);
                ComUtilities.Release(ref area);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref charts);
                ComUtilities.Release(ref sheet);
            }
        });
        var response = _fixture.Send("window.get-context", new { });
        using var result = JsonDocument.Parse(response.Result!);
        var window = Assert.Single(result.RootElement.GetProperty("windows").EnumerateArray());
        Assert.Equal("chart", window.GetProperty("selectionKind").GetString());
        Assert.Contains("ContextChart", window.GetProperty("activeChartName").GetString(), StringComparison.Ordinal);
        Assert.False(window.TryGetProperty("selectedRangeAddress", out var address) &&
            address.ValueKind != JsonValueKind.Null);
        Assert.False(window.TryGetProperty("activeCellAddress", out var cell) &&
            cell.ValueKind != JsonValueKind.Null);
    }
}

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Window")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceOwnedContextVisibilityWindowLifecycleTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Context_ReadsEveryOwnedWindowWithoutChangingItsSelection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Excel.Window? second = null;
        bool originalVisibility = false;
        try
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? sheet = null;
                Excel.Range? firstCell = null;
                Excel.Range? secondCell = null;
                try
                {
                    sheet = ComUtilities.FindSheet(context.Book, sheetName);
                    Assert.NotNull(sheet);
                    firstCell = sheet.Range["C5"];
                    sheet.Activate();
                    firstCell.Select();
                    second = context.Book.NewWindow();
                    secondCell = sheet.Range["D7"];
                    secondCell.Select();
                    originalVisibility = context.App.Visible;
                }
                finally
                {
                    ComUtilities.Release(ref secondCell);
                    ComUtilities.Release(ref firstCell);
                    ComUtilities.Release(ref sheet);
                }
            });
            var response = _fixture.Send("window.get-context", new { });
            using var result = JsonDocument.Parse(response.Result!);
            var windows = result.RootElement.GetProperty("windows");
            Assert.Equal(2, windows.GetArrayLength());
            Assert.Equal(["$C$5", "$D$7"], windows.EnumerateArray()
                .Select(item => item.GetProperty("selectedRangeAddress").GetString()).Order());
            Assert.Equal(originalVisibility, result.RootElement.GetProperty("isApplicationVisible").GetBoolean());
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Range? selected = null;
                try
                {
                    Assert.Equal(originalVisibility, context.App.Visible);
                    selected = second!.Selection as Excel.Range;
                    Assert.NotNull(selected);
                    Assert.Equal("$D$7", selected.Address);
                }
                finally
                {
                    ComUtilities.Release(ref selected);
                }
            });
        }
        finally
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                try
                {
                    second?.Close(SaveChanges: false);
                }
                finally
                {
                    ComUtilities.Release(ref second);
                }
            });
        }
    }

    [Fact]
    public void Visibility_GroupedRowsReportOutlineLevelButNotAClaimedCause()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Range? rows = null;
            object? outcome = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range["A2:A3"];
                rows = range.EntireRow;
                outcome = rows.Group();
                rows.Hidden = true;
            }
            finally
            {
                ComUtilities.Release(ref outcome);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
        var response = _fixture.Send("rangeformat.get-visibility", new
        {
            sheetName,
            rangeAddress = "A2:A3",
            axis = "rows"
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.All(result.RootElement.GetProperty("items").EnumerateArray(), item =>
        {
            Assert.True(item.GetProperty("hidden").GetBoolean());
            Assert.Equal(2, item.GetProperty("outlineLevel").GetInt32());
            Assert.Equal("undetermined", item.GetProperty("hiddenCause").GetString());
        });
    }
}

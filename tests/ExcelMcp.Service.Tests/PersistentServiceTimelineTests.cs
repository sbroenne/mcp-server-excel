using System.Text.Json;
using Xunit;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Timelines")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceTimelineTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Timeline_CreateReadAndFilterActualDates()
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-timeline", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.set-timeline-selection", new
        {
            slicerName = name,
            timelineSelection = new { startDate = new DateTime(2024, 2, 1), endDate = new DateTime(2024, 2, 29) }
        });
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        var details = state.RootElement.GetProperty("slicer");
        Assert.True(details.GetProperty("isTimeline").GetBoolean());
        Assert.False(details.GetProperty("filterCleared").GetBoolean());
        Assert.Equal("2024-02-01", details.GetProperty("timeline").GetProperty("startDate").GetDateTime().ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture));
        var data = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = pivot });
        using var values = JsonDocument.Parse(data.Result!);
        var rows = values.RootElement.GetProperty("values");
        Assert.Equal(20d, rows[rows.GetArrayLength() - 1][1].GetDouble());
        _fixture.Send("slicer.clear-timeline-selection", new { slicerName = name });
        var cleared = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var clear = JsonDocument.Parse(cleared.Result!);
        Assert.True(clear.RootElement.GetProperty("slicer").GetProperty("filterCleared").GetBoolean());
    }

    [Fact]
    public void Slicer_LayoutUpdateReadsNativeDimensionsAndCaption()
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-slicer", new
        {
            pivotTableName = pivot,
            fieldName = "Region",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.update-slicer", new
        {
            slicerName = name,
            slicerOptions = new { width = 240d, height = 180d, columnCount = 2, caption = "Areas", displayHeader = false }
        });
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        var details = state.RootElement.GetProperty("slicer");
        Assert.Equal("Areas", details.GetProperty("caption").GetString());
        Assert.Equal(240d, details.GetProperty("width").GetDouble(), 2);
        Assert.Equal(180d, details.GetProperty("height").GetDouble(), 2);
        Assert.Equal(2, details.GetProperty("columnCount").GetInt32());
        Assert.False(details.GetProperty("displayHeader").GetBoolean());
    }

    [Fact]
    public void Timeline_ListReturnsDateStateRatherThanAnOrdinaryEmptySlicer()
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-timeline", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        var read = _fixture.Send("slicer.list-slicers", new { pivotTableName = pivot });
        using var state = JsonDocument.Parse(read.Result!);
        var control = Assert.Single(state.RootElement.GetProperty("slicers").EnumerateArray(),
            item => item.GetProperty("name").GetString() == name);
        Assert.True(control.GetProperty("isTimeline").GetBoolean());
        Assert.Equal("Months", control.GetProperty("timeline").GetProperty("granularity").GetString());
    }

    [Fact]
    public void OrdinaryDateSlicer_DoesNotReuseATimelineCache()
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-timeline", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        var ordinary = name + "_Items";
        _fixture.Send("slicer.create-slicer", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = ordinary,
            destinationSheet = sheet,
            position = "J12"
        });
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = ordinary });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.False(state.RootElement.GetProperty("slicer").GetProperty("isTimeline").GetBoolean());
        Assert.Equal(3, state.RootElement.GetProperty("slicer").GetProperty("availableItems").GetArrayLength());
    }

    [Theory]
    [InlineData("Years")]
    [InlineData("Quarters")]
    [InlineData("Months")]
    [InlineData("Days")]
    public void Timeline_DisplayLevelAndViewFlagsRoundTrip(string granularity)
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-timeline", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.update-slicer", new
        {
            slicerName = name,
            slicerOptions = new
            {
                granularity,
                showHeader = false,
                showSelectionLabel = false,
                showTimeLevel = false,
                showHorizontalScrollbar = false,
                width = 350d
            }
        });
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        var timeline = state.RootElement.GetProperty("slicer").GetProperty("timeline");
        Assert.Equal(granularity, timeline.GetProperty("granularity").GetString());
        Assert.False(timeline.GetProperty("showHeader").GetBoolean());
        Assert.False(timeline.GetProperty("showSelectionLabel").GetBoolean());
        Assert.False(timeline.GetProperty("showTimeLevel").GetBoolean());
        Assert.False(timeline.GetProperty("showHorizontalScrollbar").GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedCache_ConnectAndDisconnectWithoutReplacingTheCache(bool timeline)
    {
        var (sheet, pivot, name) = CreateSource();
        var second = pivot + "_Linked";
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.PivotTable? source = null;
            Excel.PivotCache? cache = null;
            Excel.Worksheet? worksheet = null;
            Excel.PivotTables? tables = null;
            Excel.Range? destination = null;
            Excel.PivotTable? linked = null;
            try
            {
                source = CoreLookupHelpers.FindPivotTable(ctx.Book, pivot);
                cache = source.PivotCache();
                worksheet = ComUtilities.FindSheet(ctx.Book, sheet);
                tables = (Excel.PivotTables)worksheet!.PivotTables();
                destination = worksheet.Range["E15"];
                linked = tables.Add(cache, destination, second);
            }
            finally
            {
                ComUtilities.Release(ref linked);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref source);
            }
        });
        _fixture.Send("pivottablefield.add-row-field", new { pivotTableName = second, fieldName = "Region" });
        _fixture.Send("pivottablefield.add-value-field", new { pivotTableName = second, fieldName = "Amount", customName = "Total" });
        _fixture.Send(timeline ? "slicer.create-timeline" : "slicer.create-slicer", new
        {
            pivotTableName = pivot,
            fieldName = timeline ? "Date" : "Region",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.connect-pivottable", new { slicerName = name, pivotTableName = second });
        if (timeline)
            _fixture.Send("slicer.set-timeline-selection", new
            {
                slicerName = name,
                timelineSelection = new { startDate = new DateTime(2024, 2, 1), endDate = new DateTime(2024, 2, 29) }
            });
        else
            _fixture.Send("slicer.set-slicer-selection", new { slicerName = name, selectedItems = new List<string> { "North" } });
        foreach (var selectedPivot in new[] { pivot, second })
        {
            var filtered = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = selectedPivot });
            using var values = JsonDocument.Parse(filtered.Result!);
            var rows = values.RootElement.GetProperty("values");
            Assert.Equal(timeline ? 20d : 40d, rows[rows.GetArrayLength() - 1][1].GetDouble());
        }
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using (var state = JsonDocument.Parse(read.Result!))
        {
            var connections = state.RootElement.GetProperty("slicer").GetProperty("connectedPivotTables").EnumerateArray().Select(item => item.GetString()).ToArray();
            Assert.Equal(2, connections.Length);
            Assert.Contains(pivot, connections);
            Assert.Contains(second, connections);
        }
        _fixture.Send("slicer.disconnect-pivottable", new { slicerName = name, pivotTableName = second });
        var retained = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var remaining = JsonDocument.Parse(retained.Result!);
        Assert.Equal(pivot, Assert.Single(remaining.RootElement.GetProperty("slicer").GetProperty("connectedPivotTables").EnumerateArray()).GetString());
    }

    [Fact]
    public async Task IncompatibleCache_RejectsConnectionAndRetainsSource()
    {
        var (sheet, pivot, name) = CreateSource();
        var (_, other, _) = CreateSource();
        _fixture.Send("slicer.create-slicer", new
        {
            pivotTableName = pivot,
            fieldName = "Region",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        var rejected = await _fixture.SendForFailureAsync("slicer.connect-pivottable", new { slicerName = name, pivotTableName = other });
        Assert.False(rejected.Success);
        Assert.Contains("PivotCache", rejected.ErrorMessage, StringComparison.Ordinal);
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.Equal(pivot, Assert.Single(state.RootElement.GetProperty("slicer").GetProperty("connectedPivotTables").EnumerateArray()).GetString());
    }

    [Theory]
    [InlineData("""{"width":0}""")]
    [InlineData("""{"height":-1}""")]
    [InlineData("""{"columnCount":0}""")]
    [InlineData("""{"style":"missing-style"}""")]
    [InlineData("""{"granularity":"Years"}""")]
    [InlineData("""{"unknown":true}""")]
    public async Task InvalidLayout_IsRejectedBeforeChangingCaption(string options)
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-slicer", new
        {
            pivotTableName = pivot,
            fieldName = "Region",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        using var document = JsonDocument.Parse(options);
        var rejected = await _fixture.SendForFailureAsync("slicer.update-slicer", new
        {
            slicerName = name,
            slicerOptions = document.RootElement
        });
        Assert.False(rejected.Success);
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.Equal(name, state.RootElement.GetProperty("slicer").GetProperty("caption").GetString());
    }

    [Fact]
    public async Task Timeline_RejectsOrdinaryItemSelectionAndInvalidDateOrder()
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-timeline", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        var wrongOperation = await _fixture.SendForFailureAsync("slicer.set-slicer-selection", new
        {
            slicerName = name,
            selectedItems = new List<string> { "2024" }
        });
        Assert.False(wrongOperation.Success);
        Assert.Contains("set-timeline-selection", wrongOperation.ErrorMessage, StringComparison.Ordinal);
        var invalidRange = await _fixture.SendForFailureAsync("slicer.set-timeline-selection", new
        {
            slicerName = name,
            timelineSelection = new { startDate = new DateTime(2024, 3, 1), endDate = new DateTime(2024, 1, 1) }
        });
        Assert.False(invalidRange.Success);
    }

    [Theory]
    [InlineData(false, "SlicerStyleLight2")]
    [InlineData(true, "TimeSlicerStyleLight2")]
    public void NativeStyles_AreAppliedWithoutChangingCaption(bool timeline, string style)
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send(timeline ? "slicer.create-timeline" : "slicer.create-slicer", new
        {
            pivotTableName = pivot,
            fieldName = timeline ? "Date" : "Region",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.update-slicer", new { slicerName = name, slicerOptions = new { style } });
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.Equal(style, state.RootElement.GetProperty("slicer").GetProperty("style").GetString());
        Assert.Equal(name, state.RootElement.GetProperty("slicer").GetProperty("caption").GetString());
    }

    [Fact]
    public async Task Timeline_SaveReopenPreservesDateFilterAndDisplayLevel()
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-timeline", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.set-timeline-selection", new
        {
            slicerName = name,
            timelineSelection = new { startDate = new DateTime(2024, 2, 1), endDate = new DateTime(2024, 2, 29) }
        });
        _fixture.Send("slicer.update-slicer", new { slicerName = name, slicerOptions = new { granularity = "Days", caption = "Dates" } });
        await _fixture.SaveAndReopenAsync();
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        var details = state.RootElement.GetProperty("slicer");
        Assert.Equal("Dates", details.GetProperty("caption").GetString());
        Assert.Equal("Days", details.GetProperty("timeline").GetProperty("granularity").GetString());
        Assert.Equal(new DateTime(2024, 2, 1), details.GetProperty("timeline").GetProperty("startDate").GetDateTime().Date);
        Assert.Equal(new DateTime(2024, 2, 29), details.GetProperty("timeline").GetProperty("endDate").GetDateTime().Date);
    }

    [Fact]
    public async Task Timeline_DeleteRemovesOnlySelectedControl()
    {
        var (sheet, pivot, name) = CreateSource();
        _fixture.Send("slicer.create-timeline", new
        {
            pivotTableName = pivot,
            fieldName = "Date",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.delete-slicer", new { slicerName = name });
        var missing = await _fixture.SendForFailureAsync("slicer.get-slicer", new { slicerName = name });
        Assert.False(missing.Success);
        var list = _fixture.Send("slicer.list-slicers", new { pivotTableName = pivot });
        using var state = JsonDocument.Parse(list.Result!);
        Assert.Equal(0, state.RootElement.GetProperty("slicers").GetArrayLength());
    }

    [Fact]
    public async Task TableSlicer_LayoutWorksButPivotConnectionsAreRejected()
    {
        var (sheet, pivot, name) = CreateSource();
        var tableName = $"Source_{Guid.NewGuid():N}";
        _fixture.Send("table.create", new { sheetName = sheet, tableName, rangeAddress = "A1:C4" });
        _fixture.Send("slicer.create-table-slicer", new
        {
            tableName,
            columnName = "Region",
            slicerName = name,
            destinationSheet = sheet,
            position = "J2"
        });
        _fixture.Send("slicer.update-slicer", new { slicerName = name, slicerOptions = new { columnCount = 2, caption = "Areas" } });
        var read = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        using var state = JsonDocument.Parse(read.Result!);
        var details = state.RootElement.GetProperty("slicer");
        Assert.True(details.GetProperty("isTable").GetBoolean());
        Assert.Equal(tableName, details.GetProperty("connectedTable").GetString());
        Assert.Equal(2, details.GetProperty("availableItems").GetArrayLength());
        Assert.Equal(2, details.GetProperty("columnCount").GetInt32());
        var rejected = await _fixture.SendForFailureAsync("slicer.connect-pivottable", new { slicerName = name, pivotTableName = pivot });
        Assert.False(rejected.Success);
    }

    private (string Sheet, string Pivot, string Slicer) CreateSource()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var pivot = $"Dashboard_{Guid.NewGuid():N}";
        var name = $"Control_{Guid.NewGuid():N}";
        _fixture.Send("range.set-values", new
        {
            sheetName = sheet,
            rangeAddress = "A1:C4",
            values = new object[][]
            {
                ["Date", "Region", "Amount"],
                [new DateTime(2024, 1, 10).ToOADate(), "North", 10],
                [new DateTime(2024, 2, 10).ToOADate(), "South", 20],
                [new DateTime(2024, 3, 10).ToOADate(), "North", 30]
            }
        });
        _fixture.Send("range.set-number-format", new { sheetName = sheet, rangeAddress = "A2:A4", formatCode = "yyyy-mm-dd" });
        _fixture.Send("pivottable.create-from-range", new
        {
            sourceSheet = sheet,
            sourceRange = "A1:C4",
            destinationSheet = sheet,
            destinationCell = "E1",
            pivotTableName = pivot
        });
        _fixture.Send("pivottablefield.add-row-field", new { pivotTableName = pivot, fieldName = "Region" });
        _fixture.Send("pivottablefield.add-value-field", new
        {
            pivotTableName = pivot,
            fieldName = "Amount",
            aggregationFunction = "Sum",
            customName = "Total"
        });
        return (sheet, pivot, name);
    }
}

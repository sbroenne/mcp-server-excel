using System.Text.Json;
using System.Collections.Concurrent;
using System.IO.Compression;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceChartDepthTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void ErrorBars_SetAndReadActualNativePresence()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-error-bars", new { chartName, seriesIndex = 1, errorBarOptions = new { kind = "Fixed", amount = 2d } });
        var response = _fixture.Send("chartconfig.get-error-bars", new { chartName, seriesIndex = 1 });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.True(state.RootElement.GetProperty("hasErrorBars").GetBoolean());
        Assert.False(state.RootElement.GetProperty("settingsReadable").GetBoolean());
    }

    [Theory]
    [InlineData("Fixed", 2d, "Cap")]
    [InlineData("Percent", 5d, "NoCap")]
    [InlineData("StandardDeviation", 1d, "Cap")]
    [InlineData("StandardError", null, "NoCap")]
    public void ErrorBars_AllCalculatedKindsAndClearRoundTrip(string kind, double? amount, string endStyle)
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-error-bars", new { chartName, seriesIndex = 1, errorBarOptions = new { kind, amount, endStyle } });
        var response = _fixture.Send("chartconfig.get-error-bars", new { chartName, seriesIndex = 1 });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.True(state.RootElement.GetProperty("hasErrorBars").GetBoolean());
        Assert.Equal(endStyle, state.RootElement.GetProperty("endStyle").GetString());
        _fixture.Send("chartconfig.set-error-bars", new { chartName, seriesIndex = 1, errorBarOptions = new { enabled = false } });
        var cleared = _fixture.Send("chartconfig.get-error-bars", new { chartName, seriesIndex = 1 });
        using var clear = JsonDocument.Parse(cleared.Result!);
        Assert.False(clear.RootElement.GetProperty("hasErrorBars").GetBoolean());
    }

    [Theory]
    [InlineData("Both")]
    [InlineData("Plus")]
    [InlineData("Minus")]
    public void CustomErrorBars_NativeRangesAndIncludesAreAccepted(string include)
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("range.set-values", new { sheetName = sheet, rangeAddress = "E1:F3", values = new object[][] { [1d, 2d], [2d, 3d], [3d, 4d] } });
        var response = _fixture.Send("chartconfig.set-error-bars", new
        {
            chartName,
            seriesIndex = 1,
            errorBarOptions = new { kind = "Custom", plusRange = "E1:E3", minusRange = "F1:F3", include }
        });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.True(state.RootElement.GetProperty("hasErrorBars").GetBoolean());
        Assert.False(state.RootElement.GetProperty("settingsReadable").GetBoolean());
        Assert.Contains("custom source", state.RootElement.GetProperty("readLimitations").GetString(), StringComparison.OrdinalIgnoreCase);
        var xml = ReadSavedChartXml(sheet);
        XNamespace charts = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var bars = Assert.Single(xml.Descendants(charts + "errBars"));
        Assert.Equal("cust", bars.Element(charts + "errValType")!.Attribute("val")!.Value);
        Assert.Equal(include.ToLowerInvariant(), bars.Element(charts + "errBarType")!.Attribute("val")!.Value);
        if (include != "Minus")
            Assert.Contains("$E$1:$E$3", bars.Element(charts + "plus")!.Descendants(charts + "f").Single().Value, StringComparison.Ordinal);
        if (include != "Plus")
            Assert.Contains("$F$1:$F$3", bars.Element(charts + "minus")!.Descendants(charts + "f").Single().Value, StringComparison.Ordinal);
    }

    [Fact]
    public void HorizontalErrorBars_WorkForNativeScatterSeries()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-chart-type", new { chartName, chartType = "XYScatter" });
        var response = _fixture.Send("chartconfig.set-error-bars", new
        {
            chartName,
            seriesIndex = 1,
            errorBarOptions = new { direction = "X", kind = "Fixed", amount = 2d }
        });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.True(state.RootElement.GetProperty("hasErrorBars").GetBoolean());
    }

    [Fact]
    public void PointFormat_ChangesOnlySelectedNativePoint()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-point-format", new { chartName, seriesIndex = 1, pointIndex = 2, pointOptions = new { fillColor = "#FF0000", lineColor = "#0000FF", lineWeight = 2d } });
        var response = _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 1, pointIndex = 2 });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.Equal("#FF0000", state.RootElement.GetProperty("fillColor").GetString());
        Assert.Equal("#0000FF", state.RootElement.GetProperty("lineColor").GetString());
        var other = _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 1, pointIndex = 1 });
        using var unchanged = JsonDocument.Parse(other.Result!);
        Assert.NotEqual("#FF0000", unchanged.RootElement.GetProperty("fillColor").GetString());
    }

    [Fact]
    public void PointMarkers_LineSeriesRoundTripsSizeStyleAndReportsTransparencyLimit()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-series-chart-type", new { chartName, seriesIndex = 2, chartType = "LineMarkers" });
        _fixture.Send("chartconfig.set-series-axis-group", new { chartName, seriesIndex = 2, axisGroup = "Secondary" });
        var neighbor = _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 2, pointIndex = 2 }).Result;
        _fixture.Send("chartconfig.set-point-format", new
        {
            chartName,
            seriesIndex = 2,
            pointIndex = 1,
            pointOptions = new { markerStyle = "Diamond", markerSize = 14, fillColor = "#70AD47" }
        });
        var response = _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 2, pointIndex = 1 });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.True(state.RootElement.GetProperty("markersSupported").GetBoolean());
        Assert.Equal("Diamond", state.RootElement.GetProperty("markerStyle").GetString());
        Assert.Equal(14, state.RootElement.GetProperty("markerSize").GetInt32());
        Assert.Equal("#70AD47", state.RootElement.GetProperty("fillColor").GetString());
        Assert.False(state.RootElement.GetProperty("fillTransparencyAvailable").GetBoolean());
        Assert.Contains("cannot be inspected", state.RootElement.GetProperty("fillTransparencyReadError").GetString(), StringComparison.Ordinal);
        var series = _fixture.Send("chartconfig.get-series-settings", new { chartName, seriesIndex = 2 });
        using var settings = JsonDocument.Parse(series.Result!);
        Assert.Equal("LineMarkers", settings.RootElement.GetProperty("chartType").GetString());
        Assert.Equal("Secondary", settings.RootElement.GetProperty("axisGroup").GetString());
        Assert.Equal(3, settings.RootElement.GetProperty("pointCount").GetInt32());
        Assert.Equal(neighbor, _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 2, pointIndex = 2 }).Result);
    }

    [Theory]
    [InlineData("ColumnClustered")]
    public void PointTransparency_NativeWritePersistsActualAlphaDespiteGetterLimit(string chartType)
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-chart-type", new { chartName, chartType });
        _fixture.Send("chartconfig.set-point-format", new
        {
            chartName,
            seriesIndex = 1,
            pointIndex = 1,
            pointOptions = new { fillColor = "#FF0000", fillTransparency = 0.25 }
        });
        var xml = ReadSavedChartXml(sheet);
        XNamespace charts = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var point = Assert.Single(xml.Descendants(charts + "ser").First().Elements(charts + "dPt"));
        Assert.Equal("0", point.Element(charts + "idx")!.Attribute("val")!.Value);
        var alpha = Assert.Single(point.Descendants(drawing + "alpha"));
        Assert.Equal("75000", alpha.Attribute("val")!.Value);
    }

    [Fact]
    public async Task MarkerTransparency_IsRejectedBeforeChangingColorOrMarkerSettings()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-series-chart-type", new { chartName, seriesIndex = 2, chartType = "LineMarkers" });
        var before = _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 2, pointIndex = 1 }).Result;
        var response = await _fixture.SendForFailureAsync("chartconfig.set-point-format", new
        {
            chartName,
            seriesIndex = 2,
            pointIndex = 1,
            pointOptions = new { markerStyle = "Diamond", markerSize = 14, fillColor = "#FF0000", fillTransparency = 0.25 }
        });
        Assert.False(response.Success);
        Assert.Contains("transparency", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 2, pointIndex = 1 }).Result);
    }

    [Fact]
    public void ColumnPointTransparency_ReadReportsNativeGetterLimit()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-point-format", new { chartName, seriesIndex = 1, pointIndex = 1, pointOptions = new { fillColor = "#70AD47", fillTransparency = 0.25 } });
        var response = _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 1, pointIndex = 1 });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.False(state.RootElement.GetProperty("fillTransparencyAvailable").GetBoolean());
        Assert.Contains("cannot be inspected", state.RootElement.GetProperty("fillTransparencyReadError").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public void SecondaryNumberFormat_LeavesPrimaryAxisFormatUnchanged()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-series-axis-group", new { chartName, seriesIndex = 2, axisGroup = "Secondary" });
        _fixture.Send("chartconfig.set-axis-number-format", new { chartName, axis = "Value", numberFormat = "0.00" });
        _fixture.Send("chartconfig.set-axis-number-format", new { chartName, axis = "ValueSecondary", numberFormat = "0%" });
        var secondary = _fixture.Send("chartconfig.get-axis-number-format", new { chartName, axis = "ValueSecondary" });
        var primary = _fixture.Send("chartconfig.get-axis-number-format", new { chartName, axis = "Value" });
        Assert.Equal("\"0%\"", secondary.Result);
        Assert.Equal("\"0.00\"", primary.Result);
    }

    [Theory]
    [InlineData("get-series-settings", """{"seriesIndex":0}""")]
    [InlineData("set-series-axis-group", """{"seriesIndex":99,"axisGroup":"Secondary"}""")]
    [InlineData("set-series-axis-group", """{"seriesIndex":1,"axisGroup":"Unknown"}""")]
    [InlineData("set-series-chart-type", """{"seriesIndex":1,"chartType":"Unknown"}""")]
    [InlineData("set-error-bars", """{"seriesIndex":1,"errorBarOptions":{"kind":"Fixed","amount":-1}}""")]
    [InlineData("set-error-bars", """{"seriesIndex":1,"errorBarOptions":{"kind":"StandardError","amount":1}}""")]
    [InlineData("set-error-bars", """{"seriesIndex":1,"errorBarOptions":{"kind":"Custom","plusRange":"A2:A3","minusRange":"B2:B4"}}""")]
    [InlineData("set-error-bars", """{"seriesIndex":1,"errorBarOptions":{"kind":"Fixed","amount":2,"direction":"X"}}""")]
    [InlineData("set-error-bars", """{"seriesIndex":1,"errorBarOptions":{"kind":"Fixed","amount":2,"plusRange":"A2:A4"}}""")]
    [InlineData("set-error-bars", """{"seriesIndex":1,"errorBarOptions":{"kind":"Unknown","amount":2}}""")]
    [InlineData("set-point-format", """{"seriesIndex":1,"pointIndex":0,"pointOptions":{"fillColor":"#FF0000"}}""")]
    [InlineData("set-point-format", """{"seriesIndex":1,"pointIndex":99,"pointOptions":{"fillColor":"#FF0000"}}""")]
    [InlineData("set-point-format", """{"seriesIndex":1,"pointIndex":1,"pointOptions":{"fillColor":"#FF0000","lineColor":"invalid"}}""")]
    [InlineData("set-point-format", """{"seriesIndex":1,"pointIndex":1,"pointOptions":{"fillColor":"#FF0000","markerStyle":"Diamond"}}""")]
    [InlineData("set-point-format", """{"seriesIndex":1,"pointIndex":1,"pointOptions":{"fillTransparency":2}}""")]
    public async Task InvalidChartChanges_PreserveExistingSeriesAndPointState(string action, string args)
    {
        var (_, chartName) = CreateChart();
        var before = _fixture.Send("chartconfig.get-series-settings", new { chartName, seriesIndex = 1 }).Result;
        var point = _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 1, pointIndex = 1 }).Result;
        using var input = JsonDocument.Parse(args);
        var values = input.RootElement.EnumerateObject().ToDictionary(property => property.Name, property => (object?)property.Value);
        values["chartName"] = chartName;
        var response = await _fixture.SendForFailureAsync($"chartconfig.{action}", values);
        Assert.False(response.Success);
        Assert.False(string.IsNullOrEmpty(response.ErrorMessage));
        Assert.Equal(before, _fixture.Send("chartconfig.get-series-settings", new { chartName, seriesIndex = 1 }).Result);
        Assert.Equal(point, _fixture.Send("chartconfig.get-point-format", new { chartName, seriesIndex = 1, pointIndex = 1 }).Result);
    }

    [Fact]
    public void ImageExport_CreatesActualNonemptyPng()
    {
        var (_, chartName) = CreateChart();
        var path = Path.Combine(Path.GetTempPath(), $"chart-depth-{Guid.NewGuid():N}.png");
        try
        {
            _fixture.Send("chart.export-image", new { chartName, targetPath = path });
            var data = File.ReadAllBytes(path);
            Assert.True(data.Length > 1000);
            Assert.Equal((byte)137, data[0]);
            Assert.Equal("PNG", System.Text.Encoding.ASCII.GetString(data, 1, 3));
            Assert.True(System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(data.AsSpan(16, 4)) > 100);
            Assert.True(System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(data.AsSpan(20, 4)) > 100);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("Jpeg", ".jpg", 255)]
    [InlineData("Gif", ".gif", 71)]
    public void ImageExport_OtherNativeFiltersCreateRealImages(string imageFormat, string extension, int signature)
    {
        var (_, chartName) = CreateChart();
        var path = Path.Combine(Path.GetTempPath(), $"chart-depth-{Guid.NewGuid():N}{extension}");
        try
        {
            _fixture.Send("chart.export-image", new { chartName, targetPath = path, imageFormat });
            var bytes = File.ReadAllBytes(path);
            Assert.True(bytes.Length > 1000);
            Assert.Equal((byte)signature, bytes[0]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task ImageExport_RequiresExplicitOverwriteAndPreservesExistingFileOnFailure()
    {
        var (_, chartName) = CreateChart();
        var path = Path.Combine(Path.GetTempPath(), $"chart-depth-{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllText(path, "Keep existing output");
            var rejected = await _fixture.SendForFailureAsync("chart.export-image", new { chartName, targetPath = path });
            Assert.False(rejected.Success);
            Assert.Equal("Keep existing output", File.ReadAllText(path));
            var missing = await _fixture.SendForFailureAsync("chart.export-image", new { chartName = "Missing", targetPath = path, overwrite = true });
            Assert.False(missing.Success);
            Assert.Equal("Keep existing output", File.ReadAllText(path));
            _fixture.Send("chart.export-image", new { chartName, targetPath = path, overwrite = true });
            Assert.Equal((byte)137, File.ReadAllBytes(path)[0]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task ImageExport_FailedReplacementRemovesActualNativeTemporaryImage()
    {
        var (_, chartName) = CreateChart();
        var directory = Path.Combine(Path.GetTempPath(), $"chart-export-replace-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "sales.png");
        try
        {
            File.WriteAllText(path, "Keep existing output");
            var before = _fixture.Send("chartconfig.get-series-settings", new { chartName, seriesIndex = 1 }).Result;
            using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var createdImages = new ConcurrentQueue<string>();
                using var watcher = new FileSystemWatcher(directory, ".sales.*.tmp.png");
                watcher.Created += (_, change) => createdImages.Enqueue(change.FullPath);
                watcher.EnableRaisingEvents = true;
                var rejected = await _fixture.SendForFailureAsync("chart.export-image",
                    new { chartName, targetPath = path, overwrite = true });
                Assert.Equal(nameof(UnauthorizedAccessException), rejected.ExceptionType);
                Assert.True(SpinWait.SpinUntil(() => !createdImages.IsEmpty, TimeSpan.FromSeconds(5)),
                    "Excel must write its temporary image before the replacement failure.");
                Assert.Equal("Keep existing output", File.ReadAllText(path));
                Assert.Equal(new[] { path }, Directory.GetFiles(directory));
            }
            Assert.Equal(before, _fixture.Send("chartconfig.get-series-settings", new { chartName, seriesIndex = 1 }).Result);
            var response = _fixture.Send("chart.export-image", new { chartName, targetPath = path, overwrite = true });
            using var result = JsonDocument.Parse(response.Result!);
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(path, result.RootElement.GetProperty("filePath").GetString());
            Assert.Equal("export-image", result.RootElement.GetProperty("action").GetString());
            Assert.Equal((byte)137, File.ReadAllBytes(path)[0]);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("set-series-axis-group", """{"axisGroup":"Secondary"}""")]
    [InlineData("set-error-bars", """{"errorBarOptions":{"kind":"Fixed","amount":2}}""")]
    [InlineData("set-point-format", """{"pointIndex":1,"pointOptions":{"fillColor":"#FF0000"}}""")]
    public async Task PivotChart_PerSeriesMutationsAreRejectedWithoutBreakingPivotLink(string action, string args)
    {
        var (sheet, _) = CreateChart();
        var pivot = $"Pivot_{Guid.NewGuid():N}";
        var chartName = $"Linked_{Guid.NewGuid():N}";
        _fixture.Send("pivottable.create-from-range", new { sourceSheet = sheet, sourceRange = "A1:C4", destinationSheet = sheet, destinationCell = "E6", pivotTableName = pivot });
        _fixture.Send("pivottablefield.add-row-field", new { pivotTableName = pivot, fieldName = "Category" });
        _fixture.Send("pivottablefield.add-value-field", new { pivotTableName = pivot, fieldName = "First" });
        _fixture.Send("chart.create-from-pivottable", new { pivotTableName = pivot, sheetName = sheet, chartType = "ColumnClustered", chartName });
        using var input = JsonDocument.Parse(args);
        var values = input.RootElement.EnumerateObject().ToDictionary(property => property.Name, property => (object?)property.Value);
        values["chartName"] = chartName;
        values["seriesIndex"] = 1;
        var rejected = await _fixture.SendForFailureAsync($"chartconfig.{action}", values);
        Assert.False(rejected.Success);
        Assert.Contains("PivotCharts", rejected.ErrorMessage, StringComparison.Ordinal);
        var read = _fixture.Send("chart.read", new { chartName });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.True(state.RootElement.GetProperty("isPivotChart").GetBoolean());
        Assert.Equal(pivot, state.RootElement.GetProperty("linkedPivotTable").GetString());
    }

    private XDocument ReadSavedChartXml(string sheet)
    {
        var path = _fixture.ExecuteRawVerification((ctx, _) =>
        {
            ctx.Book.Save();
            return ctx.Book.FullName;
        });
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        XNamespace charts = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        return archive.Entries.Where(entry => entry.FullName.StartsWith("xl/charts/chart", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .Select(entry =>
            {
                using var stream = entry.Open();
                return XDocument.Load(stream);
            }).Single(document => document.Descendants(charts + "f").Any(node => node.Value.Contains(sheet, StringComparison.Ordinal)));
    }

    [Fact]
    public void SeriesAxisGroup_RoundTripsActualSecondaryAssignment()
    {
        var (_, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-series-axis-group", new { chartName, seriesIndex = 2, axisGroup = "Secondary" });
        var response = _fixture.Send("chartconfig.get-series-settings", new { chartName, seriesIndex = 2 });
        using var state = JsonDocument.Parse(response.Result!);
        Assert.Equal("Secondary", state.RootElement.GetProperty("axisGroup").GetString());
        var read = _fixture.Send("chart.read", new { chartName });
        Assert.True(read.Success, read.ErrorMessage);
        using var chartState = JsonDocument.Parse(read.Result!);
        var series = chartState.RootElement.GetProperty("series");
        Assert.Equal("Secondary", series[1].GetProperty("axisGroup").GetString());
        Assert.Equal("ColumnClustered", series[1].GetProperty("chartType").GetString());
        var values = series[1].GetProperty("values");
        Assert.Equal(3, values.GetArrayLength());
        Assert.Equal(100, values[0].GetInt32());
        Assert.Equal(200, values[1].GetInt32());
        Assert.Equal(300, values[2].GetInt32());
    }

    [Fact]
    public void SecondaryAxisTitle_DoesNotChangePrimaryCategoryTitle()
    {
        var (sheet, chartName) = CreateChart();
        WithChart(sheet, chartName, chart =>
        {
            Excel.SeriesCollection? series = null;
            Excel.Series? second = null;
            try
            {
                series = (Excel.SeriesCollection)chart.SeriesCollection();
                second = series.Item(2);
                second.AxisGroup = Excel.XlAxisGroup.xlSecondary;
            }
            finally
            {
                ComUtilities.Release(ref second);
                ComUtilities.Release(ref series);
            }
        });
        _fixture.Send("chartconfig.set-axis-title", new { chartName, axis = "ValueSecondary", title = "Secondary values" });
        WithChart(sheet, chartName, chart =>
        {
            Excel.Axis? secondary = null;
            Excel.Axis? primary = null;
            Excel.AxisTitle? title = null;
            try
            {
                secondary = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary);
                primary = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Assert.True(secondary.HasTitle);
                Assert.False(primary.HasTitle);
                title = secondary.AxisTitle;
                Assert.Equal("Secondary values", title.Text);
            }
            finally
            {
                ComUtilities.Release(ref title);
                ComUtilities.Release(ref primary);
                ComUtilities.Release(ref secondary);
            }
        });
    }

    [Theory]
    [InlineData("Secondary", Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary, Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary)]
    [InlineData("CategorySecondary", Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary, Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary)]
    public void AxisTitle_LegacyAndExplicitSelectors_PreserveOtherAxis(
        string axis, Excel.XlAxisType targetType, Excel.XlAxisGroup targetGroup,
        Excel.XlAxisType otherType, Excel.XlAxisGroup otherGroup)
    {
        var (sheet, chartName) = CreateChart();
        _fixture.Send("chartconfig.set-series-axis-group", new { chartName, seriesIndex = 2, axisGroup = "Secondary" });
        WithChart(sheet, chartName, chart =>
            chart.HasAxis[Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary] = true);
        _fixture.Send("chartconfig.set-axis-title", new { chartName, axis, title = "Secondary categories" });
        WithChart(sheet, chartName, chart =>
        {
            Excel.Axis? secondary = null;
            Excel.Axis? primary = null;
            Excel.AxisTitle? title = null;
            try
            {
                secondary = (Excel.Axis)chart.Axes(targetType, targetGroup);
                primary = (Excel.Axis)chart.Axes(otherType, otherGroup);
                Assert.True(secondary.HasTitle);
                Assert.False(primary.HasTitle);
                title = secondary.AxisTitle;
                Assert.Equal("Secondary categories", title.Text);
            }
            finally
            {
                ComUtilities.Release(ref title);
                ComUtilities.Release(ref primary);
                ComUtilities.Release(ref secondary);
            }
        });
    }

    private (string Sheet, string Chart) CreateChart()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Depth_{Guid.NewGuid():N}";
        _fixture.Send("range.set-values", new
        {
            sheetName = sheet,
            rangeAddress = "A1:C4",
            values = new object[][] { ["Category", "First", "Second"], ["A", 10, 100], ["B", 20, 200], ["C", 30, 300] }
        });
        _fixture.Send("chart.create-from-range", new { sheetName = sheet, sourceRangeAddress = "A1:C4", chartType = "ColumnClustered", chartName = name });
        return (sheet, name);
    }

    private void WithChart(string sheetName, string name, Action<Excel.Chart> action)
    {
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.Chart? chart = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                shapes = sheet.Shapes;
                shape = shapes.Item(name);
                chart = shape.Chart;
                action(chart);
            }
            finally
            {
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}

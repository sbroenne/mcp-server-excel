using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePivotChartTextTests(PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceDataModelFixture>
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TextFormatting_PreservesPivotDataAndLinkAfterReopen(bool dataModel)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var pivotSheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"TextPivot_{Guid.NewGuid():N}";
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        if (dataModel)
        {
            RequireSuccess(pivots.CreateFromDataModel(batch, "RegionalSalesTable", pivotSheet, "A1", pivotName));
            RequireSuccess(pivots.AddRowField(batch, pivotName, "[RegionalSalesTable].[Region]", null));
            RequireSuccess(pivots.AddValueField(batch, pivotName, "[Measures].[TotalRevenue]", AggregationFunction.Sum, null));
        }
        else
        {
            _fixture.Send("range.set-values", new
            {
                sheetName = sheet,
                rangeAddress = "A1:B4",
                values = new object[][] { ["Product", "Sales"], ["A", 10], ["B", 20], ["C", 30] }
            });
            RequireSuccess(pivots.CreateFromRange(batch, sheet, "A1:B4", pivotSheet, "A1", pivotName));
            RequireSuccess(pivots.AddRowField(batch, pivotName, "Product", null));
            RequireSuccess(pivots.AddValueField(batch, pivotName, "Sales", AggregationFunction.Sum, null));
        }
        RequireSuccess(pivots.Refresh(batch, pivotName));
        var dataBefore = JsonSerializer.Serialize(RequireSuccess(pivots.GetData(batch, pivotName)));
        var charts = _fixture.CreateCommands<IPersistentChartCommands>();
        var created = RequireSuccess(charts.CreateFromPivotTable(batch, pivotName, sheet,
            ChartType.ColumnClustered, chartName: $"PivotText_{Guid.NewGuid():N}"));
        var chartName = created.ChartName;
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Revenue", left = 16.0, top = 12.0, alignment = "Left" });
        _fixture.Send("chartconfig.set-axis-format", new { chartName, axis = "Value", lineColor = "#D0D0D0", lineWeight = 2.5, lineVisible = false });
        _fixture.Send("chartconfig.set-area-format", new { chartName, area = "Chart", lineVisible = false });
        _fixture.Send("chartconfig.set-axis-title", new { chartName, axis = "Value", title = "Amount" });
        _fixture.Send("chartconfig.show-legend", new { chartName, visible = true });
        _fixture.Send("chartconfig.set-data-labels", new { chartName, showValue = true });
        var plottedBefore = JsonSerializer.Serialize(RequireSuccess(charts.Read(batch, chartName)).Series);
        var expected = new Dictionary<ChartTextElement, string>();
        foreach (var element in Enum.GetValues<ChartTextElement>())
        {
            var axis = element == ChartTextElement.AxisTitle ? "Value" : null;
            using var written = JsonDocument.Parse(_fixture.Send("chartconfig.set-text-format", new
            {
                chartName,
                element = element.ToString(),
                axis,
                fontName = "Arial",
                fontSize = 13.5,
                bold = true,
                italic = false,
                color = "#2468AC"
            }).Result!);
            foreach (var format in written.RootElement.GetProperty("formats").EnumerateArray())
            {
                Assert.Equal("Arial", format.GetProperty("fontName").GetString());
                Assert.Equal(13.5, format.GetProperty("fontSize").GetDouble());
                Assert.True(format.GetProperty("bold").GetBoolean());
                Assert.False(format.GetProperty("italic").GetBoolean());
                Assert.Equal("#2468AC", format.GetProperty("color").GetString());
            }
        }
        // ChartArea changes can affect inheriting text, so capture every final target before saving.
        foreach (var element in Enum.GetValues<ChartTextElement>())
            expected[element] = _fixture.Send("chartconfig.get-text-format", new
            {
                chartName,
                element = element.ToString(),
                axis = element == ChartTextElement.AxisTitle ? "Value" : null
            }).Result!;
        VerifyLayout(sheet, chartName);
        await _fixture.SaveAndReopenAsync();
        VerifyLayout(sheet, chartName);
        foreach (var (element, result) in expected)
            Assert.Equal(result, _fixture.Send("chartconfig.get-text-format", new
            {
                chartName,
                element = element.ToString(),
                axis = element == ChartTextElement.AxisTitle ? "Value" : null
            }).Result);
        var read = RequireSuccess(charts.Read(batch, chartName));
        Assert.True(read.IsPivotChart);
        Assert.Equal(pivotName, read.LinkedPivotTable);
        Assert.Equal(plottedBefore, JsonSerializer.Serialize(read.Series));
        Assert.Equal(dataBefore, JsonSerializer.Serialize(RequireSuccess(pivots.GetData(batch, pivotName))));
    }

    private void VerifyLayout(string sheetName, string chartName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.ChartTitle? title = null;
            Excel.ChartFormat? titleFormat = null;
            Excel.TextFrame2? frame = null;
            Office.TextRange2? text = null;
            Office.ParagraphFormat2? paragraph = null;
            Excel.Axis? axis = null;
            Excel.ChartFormat? axisFormat = null;
            Excel.LineFormat? axisLine = null;
            Excel.ColorFormat? axisColor = null;
            Excel.ChartArea? area = null;
            Excel.ChartFormat? areaFormat = null;
            Excel.LineFormat? areaLine = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                chartObject = (Excel.ChartObject)sheet.ChartObjects(chartName);
                chart = chartObject.Chart;
                title = chart.ChartTitle;
                Assert.Equal("Revenue", title.Text);
                Assert.Equal(16, title.Left, precision: 1);
                Assert.Equal(12, title.Top, precision: 1);
                titleFormat = title.Format;
                frame = titleFormat.TextFrame2;
                text = frame.TextRange;
                paragraph = text.ParagraphFormat;
                Assert.Equal(Office.MsoParagraphAlignment.msoAlignLeft, paragraph.Alignment);
                axis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
                axisFormat = axis.Format;
                axisLine = axisFormat.Line;
                axisColor = axisLine.ForeColor;
                Assert.Equal(0xD0D0D0, axisColor.RGB);
                Assert.Equal(2.5f, axisLine.Weight);
                Assert.Equal(Office.MsoTriState.msoFalse, axisLine.Visible);
                area = chart.ChartArea;
                areaFormat = area.Format;
                areaLine = areaFormat.Line;
                Assert.Equal(Office.MsoTriState.msoFalse, areaLine.Visible);
            }
            finally
            {
                ComUtilities.Release(ref areaLine);
                ComUtilities.Release(ref areaFormat);
                ComUtilities.Release(ref area);
                ComUtilities.Release(ref axisColor);
                ComUtilities.Release(ref axisLine);
                ComUtilities.Release(ref axisFormat);
                ComUtilities.Release(ref axis);
                ComUtilities.Release(ref paragraph);
                ComUtilities.Release(ref text);
                ComUtilities.Release(ref frame);
                ComUtilities.Release(ref titleFormat);
                ComUtilities.Release(ref title);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}

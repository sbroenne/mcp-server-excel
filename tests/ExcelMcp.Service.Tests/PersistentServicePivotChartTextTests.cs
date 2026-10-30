using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

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
        _fixture.Send("chartconfig.set-title", new { chartName, title = "Revenue" });
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
        await _fixture.SaveAndReopenAsync();
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
}

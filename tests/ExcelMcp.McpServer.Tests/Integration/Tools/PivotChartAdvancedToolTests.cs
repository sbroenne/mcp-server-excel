using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end MCP coverage for advanced PivotTable and chart operations.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PivotTables")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public sealed class PivotChartAdvancedToolTests : McpIntegrationTestBase
{
    private static readonly string[] GroupedRegions = ["North", "South"];
    private readonly string _tempDir;

    public PivotChartAdvancedToolTests(ITestOutputHelper output)
        : base(output, "PivotChartAdvancedToolClient")
    {
        _tempDir = CreateTempDirectory("PivotChartAdvancedToolTests");
    }

    [Fact]
    public async Task PivotTableAdvancedActions_ExecuteThroughMcpProtocol()
    {
        var workbookPath = Path.Join(_tempDir, $"AdvancedPivot_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await SeedPivotDataAsync(sessionId);

        AssertSuccess(await CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "create-from-range",
            ["session_id"] = sessionId,
            ["source_sheet"] = "Sheet1",
            ["source_range"] = "A1:D7",
            ["destination_sheet"] = "Sheet1",
            ["destination_cell"] = "F1",
            ["pivot_table_name"] = "AdvancedPivot"
        }), "Create advanced PivotTable");

        AssertSuccess(await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "add-row-field",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["field_name"] = "Region"
        }), "Add PivotTable row field");

        AssertSuccess(await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "add-value-field",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["field_name"] = "Sales",
            ["aggregation_function"] = "Sum"
        }), "Add PivotTable value field");

        AssertSuccess(await CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "set-cache-options",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["refresh_on_file_open"] = true,
            ["missing_items_limit"] = "None",
            ["save_source_data"] = false
        }), "Set PivotCache options");

        var cacheOptions = await CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "get-cache-options",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "AdvancedPivot"
        });
        AssertSuccess(cacheOptions, "Get PivotCache options");
        using (var cacheJson = JsonDocument.Parse(cacheOptions))
        {
            Assert.True(cacheJson.RootElement.GetProperty("refreshOnFileOpen").GetBoolean());
            Assert.Equal("None", cacheJson.RootElement.GetProperty("missingItemsLimit").GetString());
            Assert.False(cacheJson.RootElement.GetProperty("saveSourceData").GetBoolean());
        }

        var groupResult = await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "group-items",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["field_name"] = "Region",
            ["item_names"] = JsonSerializer.Serialize(GroupedRegions),
            ["group_name"] = "Core Regions"
        });
        AssertSuccess(groupResult, "Group PivotTable items");

        string groupedFieldName;
        using (var groupJson = JsonDocument.Parse(groupResult))
        {
            groupedFieldName = groupJson.RootElement.GetProperty("groupedFieldName").GetString()!;
            Assert.False(string.IsNullOrWhiteSpace(groupedFieldName));
        }

        AssertSuccess(await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "ungroup-field",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["grouped_field_name"] = groupedFieldName
        }), "Ungroup PivotTable field");

        var drillResult = await CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "drill-through",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["cell_address"] = "G2"
        });
        AssertSuccess(drillResult, "Drill through PivotTable value cell");
        using (var drillJson = JsonDocument.Parse(drillResult))
        {
            Assert.True(drillJson.RootElement.GetProperty("detailRowCount").GetInt32() > 1);
        }

        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task ChartAdvancedActions_ExecuteThroughMcpProtocol()
    {
        var workbookPath = Path.Join(_tempDir, $"AdvancedChart_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await SeedPivotDataAsync(sessionId);

        AssertSuccess(await CallToolAsync("chart", new Dictionary<string, object?>
        {
            ["action"] = "create-from-range",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["source_range_address"] = "A1:C7",
            ["chart_type"] = "ColumnClustered",
            ["chart_name"] = "AdvancedChart"
        }), "Create advanced chart");

        AssertSuccess(await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-series-chart-type",
            ["session_id"] = sessionId,
            ["chart_name"] = "AdvancedChart",
            ["series_index"] = 2,
            ["chart_type"] = "LineMarkers"
        }), "Set chart series type");

        AssertSuccess(await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-plot-options",
            ["session_id"] = sessionId,
            ["chart_name"] = "AdvancedChart",
            ["plot_by"] = "Rows",
            ["display_blanks_as"] = "Zero",
            ["plot_visible_only"] = false
        }), "Set chart plot options");

        var plotOptions = await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "get-plot-options",
            ["session_id"] = sessionId,
            ["chart_name"] = "AdvancedChart"
        });
        AssertSuccess(plotOptions, "Get chart plot options");
        Assert.Contains("\"plotBy\":\"Rows\"", plotOptions, StringComparison.Ordinal);
        Assert.Contains("\"displayBlanksAs\":\"Zero\"", plotOptions, StringComparison.Ordinal);

        AssertSuccess(await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-placement",
            ["session_id"] = sessionId,
            ["chart_name"] = "AdvancedChart",
            ["placement"] = 2,
            ["print_object"] = false,
            ["locked"] = false,
            ["rounded_corners"] = true
        }), "Set embedded chart object properties");

        AssertSuccess(await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-area-format",
            ["session_id"] = sessionId,
            ["chart_name"] = "AdvancedChart",
            ["area"] = "Chart",
            ["fill_color"] = "#FF0000",
            ["fill_transparency"] = 0.25,
            ["line_color"] = "#0000FF",
            ["line_weight"] = 2.5
        }), "Format chart area");

        AssertSuccess(await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-series-format",
            ["session_id"] = sessionId,
            ["chart_name"] = "AdvancedChart",
            ["series_index"] = 1,
            ["fill_color"] = "#00FF00",
            ["fill_transparency"] = 0.4,
            ["line_color"] = "#FF00FF",
            ["line_weight"] = 3
        }), "Format chart series material");

        await CloseSessionAsync(sessionId, save: false);
    }

    private async Task SeedPivotDataAsync(string sessionId)
    {
        var values = new object?[][]
        {
            ["Region", "Sales", "Profit", "Product"],
            ["North", 100, 30, "Widget"],
            ["North", 150, 45, "Gadget"],
            ["South", 200, 60, "Widget"],
            ["South", 125, 35, "Gadget"],
            ["West", 175, 55, "Widget"],
            ["West", 225, 70, "Gadget"]
        };

        AssertSuccess(await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:D7",
            ["values"] = values
        }), "Seed PivotTable and chart data");
    }
}

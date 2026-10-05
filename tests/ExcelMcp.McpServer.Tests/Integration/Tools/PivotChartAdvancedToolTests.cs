using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PivotTables")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "false")]
public sealed class PivotChartAdvancedToolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task PivotTableAdvancedActions_ExecuteThroughMcpProtocol()
    {
        await AssertSuccessAsync("pivottable", new()
        {
            ["action"] = "create-from-range",
            ["workbook_session_id"] = SessionId,
            ["source_sheet"] = "Sheet1",
            ["source_range"] = "A1:D7",
            ["destination_sheet"] = "Sheet1",
            ["destination_cell"] = "F1",
            ["pivot_table_name"] = "AdvancedPivot"
        }, "pivottable.create-from-range",
        """{"sourceSheet":"Sheet1","sourceRange":"A1:D7","destinationSheet":"Sheet1","destinationCell":"F1","pivotTableName":"AdvancedPivot"}""", args =>
        {
            Assert.Equal("A1:D7", args.GetProperty("sourceRange").GetString());
            Assert.Equal("F1", args.GetProperty("destinationCell").GetString());
            Assert.Equal("AdvancedPivot", args.GetProperty("pivotTableName").GetString());
        });
        await AssertSuccessAsync("pivottable_field", new()
        {
            ["action"] = "add-row-field",
            ["workbook_session_id"] = SessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["field_name"] = "Region"
        }, "pivottablefield.add-row-field",
        """{"pivotTableName":"AdvancedPivot","fieldName":"Region"}""");
        await AssertSuccessAsync("pivottable_field", new()
        {
            ["action"] = "add-value-field",
            ["workbook_session_id"] = SessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["field_name"] = "Sales",
            ["aggregation_function"] = "Sum"
        }, "pivottablefield.add-value-field",
        """{"pivotTableName":"AdvancedPivot","fieldName":"Sales","aggregationFunction":"Sum"}""", args =>
            Assert.Equal(
                "Sum",
                args.GetProperty("aggregationFunction").GetString()));
        await AssertSuccessAsync("pivottable", new()
        {
            ["action"] = "set-cache-options",
            ["workbook_session_id"] = SessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["refresh_on_file_open"] = true,
            ["missing_items_limit"] = "None",
            ["save_source_data"] = false
        }, "pivottable.set-cache-options",
        """{"pivotTableName":"AdvancedPivot","refreshOnFileOpen":true,"missingItemsLimit":"None","saveSourceData":false}""", args =>
        {
            Assert.True(args.GetProperty("refreshOnFileOpen").GetBoolean());
            Assert.Equal("None", args.GetProperty("missingItemsLimit").GetString());
            Assert.False(args.GetProperty("saveSourceData").GetBoolean());
        });

        var cacheJson = await CallAsync(
            "pivottable_read",
            new()
            {
                ["action"] = "get-cache-options",
                ["workbook_session_id"] = SessionId,
                ["pivot_table_name"] = "AdvancedPivot"
            },
            "pivottable.get-cache-options",
            """{"pivotTableName":"AdvancedPivot"}""",
            """{"success":true,"refreshOnFileOpen":true,"missingItemsLimit":"None","saveSourceData":false}""");
        using (var cache = JsonDocument.Parse(cacheJson))
        {
            Assert.True(cache.RootElement.GetProperty("refreshOnFileOpen").GetBoolean());
            Assert.Equal(
                "None",
                cache.RootElement.GetProperty("missingItemsLimit").GetString());
            Assert.False(cache.RootElement.GetProperty("saveSourceData").GetBoolean());
        }

        var groupJson = await CallAsync(
            "pivottable_field",
            new()
            {
                ["action"] = "group-items",
                ["workbook_session_id"] = SessionId,
                ["pivot_table_name"] = "AdvancedPivot",
                ["field_name"] = "Region",
                ["item_names"] = """["North","South"]""",
                ["group_name"] = "Core Regions"
            },
            "pivottablefield.group-items",
            """{"pivotTableName":"AdvancedPivot","fieldName":"Region","itemNames":["North","South"],"groupName":"Core Regions"}""",
            """{"success":true,"groupedFieldName":"Region2"}""",
            args =>
            {
                Assert.Equal(
                    ["North", "South"],
                    args.GetProperty("itemNames")
                        .EnumerateArray()
                        .Select(item => item.GetString()!)
                        .ToArray());
                Assert.Equal(
                    "Core Regions",
                    args.GetProperty("groupName").GetString());
            });
        using (var group = JsonDocument.Parse(groupJson))
        {
            Assert.Equal(
                "Region2",
                group.RootElement.GetProperty("groupedFieldName").GetString());
        }

        await AssertSuccessAsync("pivottable_field", new()
        {
            ["action"] = "ungroup-field",
            ["workbook_session_id"] = SessionId,
            ["pivot_table_name"] = "AdvancedPivot",
            ["grouped_field_name"] = "Region2"
        }, "pivottablefield.ungroup-field",
        """{"pivotTableName":"AdvancedPivot","groupedFieldName":"Region2"}""", args =>
            Assert.Equal(
                "Region2",
                args.GetProperty("groupedFieldName").GetString()));

        var drillJson = await CallAsync(
            "pivottable",
            new()
            {
                ["action"] = "drill-through",
                ["workbook_session_id"] = SessionId,
                ["pivot_table_name"] = "AdvancedPivot",
                ["cell_address"] = "G2"
            },
            "pivottable.drill-through",
            """{"pivotTableName":"AdvancedPivot","cellAddress":"G2"}""",
            """{"success":true,"detailRowCount":3}""",
            args => Assert.Equal(
                "G2",
                args.GetProperty("cellAddress").GetString()));
        using var drill = JsonDocument.Parse(drillJson);
        Assert.True(drill.RootElement.GetProperty("detailRowCount").GetInt32() > 1);
    }

    [Fact]
    public async Task ChartAdvancedActions_ExecuteThroughMcpProtocol()
    {
        await AssertSuccessAsync("chart", new()
        {
            ["action"] = "create-from-range",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Sheet1",
            ["source_range_address"] = "A1:C7",
            ["chart_type"] = "ColumnClustered",
            ["chart_name"] = "AdvancedChart"
        }, "chart.create-from-range",
        """{"sheetName":"Sheet1","sourceRangeAddress":"A1:C7","chartType":"ColumnClustered","chartName":"AdvancedChart"}""", args =>
        {
            Assert.Equal(
                "A1:C7",
                args.GetProperty("sourceRangeAddress").GetString());
            Assert.Equal(
                "ColumnClustered",
                args.GetProperty("chartType").GetString());
        });
        await AssertSuccessAsync("chart_config", new()
        {
            ["action"] = "set-series-chart-type",
            ["workbook_session_id"] = SessionId,
            ["chart_name"] = "AdvancedChart",
            ["series_index"] = 2,
            ["chart_type"] = "LineMarkers"
        }, "chartconfig.set-series-chart-type",
        """{"chartName":"AdvancedChart","seriesIndex":2,"chartType":"LineMarkers"}""", args =>
        {
            Assert.Equal(2, args.GetProperty("seriesIndex").GetInt32());
            Assert.Equal("LineMarkers", args.GetProperty("chartType").GetString());
        });
        await AssertSuccessAsync("chart_config", new()
        {
            ["action"] = "set-plot-options",
            ["workbook_session_id"] = SessionId,
            ["chart_name"] = "AdvancedChart",
            ["plot_by"] = "Rows",
            ["display_blanks_as"] = "Zero",
            ["plot_visible_only"] = false
        }, "chartconfig.set-plot-options",
        """{"chartName":"AdvancedChart","plotBy":"Rows","displayBlanksAs":"Zero","plotVisibleOnly":false}""", args =>
        {
            Assert.Equal("Rows", args.GetProperty("plotBy").GetString());
            Assert.Equal("Zero", args.GetProperty("displayBlanksAs").GetString());
            Assert.False(args.GetProperty("plotVisibleOnly").GetBoolean());
        });

        var plotJson = await CallAsync(
            "chart_config_read",
            new()
            {
                ["action"] = "get-plot-options",
                ["workbook_session_id"] = SessionId,
                ["chart_name"] = "AdvancedChart"
            },
            "chartconfig.get-plot-options",
            """{"chartName":"AdvancedChart"}""",
            """{"success":true,"plotBy":"Rows","displayBlanksAs":"Zero"}""");
        using (var plot = JsonDocument.Parse(plotJson))
        {
            Assert.Equal("Rows", plot.RootElement.GetProperty("plotBy").GetString());
            Assert.Equal(
                "Zero",
                plot.RootElement.GetProperty("displayBlanksAs").GetString());
        }

        await AssertSuccessAsync("chart_config", new()
        {
            ["action"] = "set-placement",
            ["workbook_session_id"] = SessionId,
            ["chart_name"] = "AdvancedChart",
            ["placement"] = 2,
            ["print_object"] = false,
            ["locked"] = false,
            ["rounded_corners"] = true
        }, "chartconfig.set-placement",
        """{"chartName":"AdvancedChart","placement":2,"printObject":false,"locked":false,"roundedCorners":true}""", args =>
        {
            Assert.Equal(2, args.GetProperty("placement").GetInt32());
            Assert.False(args.GetProperty("printObject").GetBoolean());
            Assert.False(args.GetProperty("locked").GetBoolean());
            Assert.True(args.GetProperty("roundedCorners").GetBoolean());
        });
        await AssertSuccessAsync("chart_config", new()
        {
            ["action"] = "set-area-format",
            ["workbook_session_id"] = SessionId,
            ["chart_name"] = "AdvancedChart",
            ["area"] = "Chart",
            ["fill_color"] = "#FF0000",
            ["fill_transparency"] = 0.25,
            ["line_color"] = "#0000FF",
            ["line_weight"] = 2.5
        }, "chartconfig.set-area-format",
        """{"chartName":"AdvancedChart","area":"Chart","fillColor":"#FF0000","fillTransparency":0.25,"lineColor":"#0000FF","lineWeight":2.5}""", args =>
        {
            Assert.Equal("#FF0000", args.GetProperty("fillColor").GetString());
            Assert.Equal(0.25, args.GetProperty("fillTransparency").GetDouble());
            Assert.Equal(2.5, args.GetProperty("lineWeight").GetDouble());
        });
        await AssertSuccessAsync("chart_config", new()
        {
            ["action"] = "set-series-format",
            ["workbook_session_id"] = SessionId,
            ["chart_name"] = "AdvancedChart",
            ["series_index"] = 1,
            ["fill_color"] = "#00FF00",
            ["fill_transparency"] = 0.4,
            ["line_color"] = "#FF00FF",
            ["line_weight"] = 3
        }, "chartconfig.set-series-format",
        """{"chartName":"AdvancedChart","seriesIndex":1,"fillColor":"#00FF00","fillTransparency":0.4,"lineColor":"#FF00FF","lineWeight":3}""", args =>
        {
            Assert.Equal(1, args.GetProperty("seriesIndex").GetInt32());
            Assert.Equal("#00FF00", args.GetProperty("fillColor").GetString());
            Assert.Equal("#FF00FF", args.GetProperty("lineColor").GetString());
        });
    }

    private async Task AssertSuccessAsync(
        string tool,
        Dictionary<string, object?> arguments,
        string command,
        string? expectedArgsJson,
        Action<JsonElement>? assertArgs = null)
    {
        var json = await CallAsync(
            tool,
            arguments,
            command,
            expectedArgsJson,
            """{"success":true}""",
            assertArgs);
        using var result = JsonDocument.Parse(json);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
    }

    private async Task<string> CallAsync(
        string tool,
        Dictionary<string, object?> arguments,
        string command,
        string? expectedArgsJson,
        string responseJson,
        Action<JsonElement>? assertArgs = null)
    {
        var call = await _fixture.CallToolAsync(
            tool,
            arguments,
            RecordingToolTest.Success(responseJson),
            command,
            expectedArgsJson);
        using var args = RecordingToolTest.ParseArgs(
            call.Request,
            command,
            SessionId);
        assertArgs?.Invoke(args.RootElement);
        return call.JsonResult;
    }
}

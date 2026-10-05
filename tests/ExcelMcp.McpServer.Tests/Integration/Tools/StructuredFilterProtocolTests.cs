using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Filtering;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "StructuredFilters")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class StructuredFilterProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ApplyFilter_PreservesNativeOptionsAndCanonicalNames(bool table)
    {
        var options = new FilterOptions { FilterOperator = FilterOperator.And, Criteria1 = ">=20", Criteria2 = "<=40" };
        var args = new Dictionary<string, object?>
        {
            ["action"] = "apply-filter",
            ["workbook_session_id"] = "session-1",
            [table ? "options" : "filter_options"] = new { filterOperator = "And", criteria1 = ">=20", criteria2 = "<=40" }
        };
        var expected = new Dictionary<string, object?> { [table ? "options" : "filterOptions"] = options };
        if (table)
        {
            args["table_name"] = "Sales";
            args["column_name"] = "Amount";
            expected["tableName"] = "Sales";
            expected["columnName"] = "Amount";
        }
        else
        {
            args["sheet_name"] = "Data";
            args["range_address"] = "A1:B6";
            args["column_index"] = 2;
            expected["sheetName"] = "Data";
            expected["rangeAddress"] = "A1:B6";
            expected["columnIndex"] = 2;
        }
        var call = await fixture.CallToolAsync(table ? "table_column" : "range_edit", args,
            RecordingToolTest.Success("""{"success":true}"""),
            table ? "tablecolumn.apply-filter" : "rangeedit.apply-filter",
            JsonSerializer.Serialize(expected, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_SeparatesFilterOptionsFromTextParsingOptions()
    {
        var tools = await fixture.ListToolsAsync();
        var range = Assert.Single(tools, tool => tool.Name == "range_edit");
        var properties = range.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("filter_options", out _));
        Assert.True(properties.TryGetProperty("options", out _));
        Assert.True(properties.TryGetProperty("clear_advanced", out _));
        Assert.True(properties.TryGetProperty("criteria_range", out _));
        Assert.True(properties.TryGetProperty("copy_to_range", out _));
        Assert.Contains("advanced row filter must be explicitly cleared before apply-filter",
            range.Description, StringComparison.Ordinal);
        Assert.Contains("do not retry or clear it without authorization", range.Description, StringComparison.Ordinal);
        var table = Assert.Single(tools, tool => tool.Name == "table_column");
        Assert.Contains("apply-filter-values is removed", table.Description, StringComparison.Ordinal);
    }
}

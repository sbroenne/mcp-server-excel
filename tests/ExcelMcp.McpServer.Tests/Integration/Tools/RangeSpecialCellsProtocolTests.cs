using System.ComponentModel;
using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeSpecialCellsProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("formulas")]
    [InlineData("constants")]
    [InlineData("blanks")]
    [InlineData("errors")]
    [InlineData("visible")]
    public async Task SpecialCells_MapsSelectorAndPreservesCompleteResult(string cellKind)
    {
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            sheetName = "Sheet1",
            rangeAddress = "$A$1:$A$64",
            cellKind,
            cellCount = 32,
            areas = Enumerable.Range(0, 32).Select(index => $"$A${index * 2 + 1}").ToArray()
        });
        var expectedArgs = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1:A64",
            cellKind
        }, ServiceProtocol.JsonOptions);

        var call = await fixture.CallToolAsync("range", new()
        {
            ["action"] = "get-special-cells",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:A64",
            ["cell_kind"] = cellKind
        }, RecordingToolTest.Success(response), "range.get-special-cells", expectedArgs);

        Assert.False(call.Result.IsError);
        using var output = JsonDocument.Parse(call.JsonResult);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(32, output.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(32, output.RootElement.GetProperty("areas").GetArrayLength());
    }

    [Fact]
    public async Task Discovery_AdvertisesRequiredSelectorAndCompleteScope()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range");
        Assert.Contains("get-special-cells", tool.Description, StringComparison.Ordinal);
        Assert.True(tool.JsonSchema.GetProperty("properties").TryGetProperty("cell_kind", out _));
        var parameter = GeneratedToolContract.GetParameter("range", "cell_kind");
        Assert.Contains("formulas", parameter.GetCustomAttribute<DescriptionAttribute>()?.Description,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("unknown-kind")]
    [InlineData("99")]
    public async Task SpecialCells_InvalidSelectorDoesNotDispatch(string cellKind)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range", new()
        {
            ["action"] = "get-special-cells",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["cell_kind"] = cellKind
        });

        Assert.True(result.IsError);
    }

    [Fact]
    public async Task SpecialCells_MissingSelectorDoesNotDispatch()
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range", new()
        {
            ["action"] = "get-special-cells",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        });

        Assert.True(result.IsError);
    }
}

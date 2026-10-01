using System.Text.Json;
using System.ComponentModel;
using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFindProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData(null)]
    [InlineData(1)]
    [InlineData(25)]
    [InlineData(1001)]
    [InlineData(int.MaxValue)]
    public async Task Find_MapsLimitAndPreservesCoverage(int? maxMatches)
    {
        var arguments = FindArguments();
        if (maxMatches.HasValue)
        {
            arguments["max_matches"] = maxMatches.Value;
        }
        var returnedCount = Math.Min(maxMatches ?? 10, 25);
        var response = new RangeFindResult
        {
            Success = true,
            TotalCount = 25,
            MatchingCells = Enumerable.Range(1, returnedCount)
                .Select(row => new RangeCell { Address = $"$A${row}", Row = row, Column = 1, Value = "Apple" })
                .ToList()
        };
        var expectedArguments = new Dictionary<string, object?>
        {
            ["sheetName"] = "Sheet1",
            ["rangeAddress"] = "A1:A26",
            ["searchValue"] = "Apple",
            ["findOptions"] = new FindOptions { MatchEntireCell = true }
        };
        if (maxMatches.HasValue)
        {
            expectedArguments["maxMatches"] = maxMatches.Value;
        }
        var expectedArgs = JsonSerializer.Serialize(expectedArguments, ServiceProtocol.JsonOptions);

        var call = await fixture.CallToolAsync(
            "range_edit", arguments,
            RecordingToolTest.Success(JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions)),
            "rangeedit.find", expectedArgs);

        using var document = JsonDocument.Parse(call.JsonResult);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal(25, root.GetProperty("totalCount").GetInt64());
        Assert.Equal(returnedCount, root.GetProperty("returnedCount").GetInt32());
        Assert.Equal(25 > returnedCount, root.GetProperty("truncated").GetBoolean());
        Assert.Equal(returnedCount, root.GetProperty("matchingCells").GetArrayLength());
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesFindLimitAndCoverage()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range_edit");
        Assert.Contains("totalCount", tool.Description, StringComparison.Ordinal);
        Assert.Contains("returnedCount", tool.Description, StringComparison.Ordinal);
        Assert.Contains("truncated", tool.Description, StringComparison.Ordinal);
        Assert.Contains("default: 10", tool.Description, StringComparison.Ordinal);
        Assert.Contains("not search time", tool.Description, StringComparison.Ordinal);
        Assert.True(tool.JsonSchema.GetProperty("properties").TryGetProperty("max_matches", out _));
        var parameter = GeneratedToolContract.GetParameter("range_edit", "max_matches");
        Assert.True(parameter.IsOptional);
        Assert.Equal(typeof(int?), parameter.ParameterType);
        Assert.Null(parameter.DefaultValue);
        Assert.Contains("2147483647", parameter.GetCustomAttribute<DescriptionAttribute>()?.Description, StringComparison.Ordinal);
        var outputProperties = Assert.IsType<JsonElement>(tool.ReturnJsonSchema).GetProperty("properties");
        Assert.True(outputProperties.TryGetProperty("totalCount", out _));
        Assert.True(outputProperties.TryGetProperty("returnedCount", out _));
        Assert.True(outputProperties.TryGetProperty("truncated", out _));
    }

    [Theory]
    [InlineData("1.5")]
    [InlineData("2147483648")]
    public async Task Find_RejectsNonIntegerOrOverflowLimitBeforeDispatch(string limitJson)
    {
        var arguments = FindArguments();
        arguments["max_matches"] = JsonSerializer.Deserialize<JsonElement>(limitJson);

        var result = await fixture.CallResultWithoutDispatchAsync("range_edit", arguments);

        Assert.True(result.IsError);
        var text = Assert.IsType<ModelContextProtocol.Protocol.TextContentBlock>(Assert.Single(result.Content)).Text;
        Assert.False(string.IsNullOrWhiteSpace(text));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public async Task Find_PreservesNonpositiveLimitAndReportsServiceError(int maxMatches)
    {
        var arguments = FindArguments();
        arguments["max_matches"] = maxMatches;
        var expectedArgs = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1:A26",
            searchValue = "Apple",
            findOptions = new FindOptions { MatchEntireCell = true },
            maxMatches
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_edit", arguments, new ServiceResponse
        {
            Success = false,
            ErrorMessage = "maxMatches must be positive.",
            ErrorCategory = "InvalidInput",
            ExceptionType = nameof(ArgumentOutOfRangeException)
        }, "rangeedit.find", expectedArgs);

        Assert.True(call.Result.IsError);
        using var document = JsonDocument.Parse(call.JsonResult);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("maxMatches", document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Sort_RejectsFindOnlyLimitBeforeDispatch()
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range_edit", new()
        {
            ["action"] = "sort",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["sort_columns"] = new[] { new { columnIndex = 1, ascending = true } },
            ["max_matches"] = 1
        });

        Assert.True(result.IsError);
        var text = Assert.IsType<ModelContextProtocol.Protocol.TextContentBlock>(Assert.Single(result.Content)).Text;
        using var document = JsonDocument.Parse(text);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("maxMatches", document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    private static Dictionary<string, object?> FindArguments() => new()
    {
        ["action"] = "find",
        ["session_id"] = "session-1",
        ["sheet_name"] = "Sheet1",
        ["range_address"] = "A1:A26",
        ["search_value"] = "Apple",
        ["find_options"] = new { matchEntireCell = true }
    };
}

using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SlicerSelectionProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData(null, "[\"2026 Q2\"]")]
    [InlineData(false, "[\"2026 Q1\"]")]
    [InlineData(true, "[]")]
    public async Task Selection_MapsCaptionsClearFirstAndResult(bool? clearFirst, string selectedItems)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "set-slicer-selection",
            ["session_id"] = "session-1",
            ["slicer_name"] = "QuarterSlicer",
            ["selected_items"] = selectedItems
        };
        var expected = new Dictionary<string, object?>
        {
            ["slicerName"] = "QuarterSlicer",
            ["selectedItems"] = JsonSerializer.Deserialize<string[]>(selectedItems)
        };
        if (clearFirst.HasValue)
        {
            arguments["clear_first"] = clearFirst.Value;
            expected["clearFirst"] = clearFirst.Value;
        }
        var call = await fixture.CallToolAsync("slicer", arguments,
            RecordingToolTest.Success("""{"success":true,"availableItems":["2026 Q1","2026 Q2"],"selectedItems":["2026 Q2"],"connectedPivotTables":["RevenuePivot"]}"""),
            "slicer.set-slicer-selection", JsonSerializer.Serialize(expected));

        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal("2026 Q2", result.RootElement.GetProperty("selectedItems")[0].GetString());
        Assert.Equal(2, result.RootElement.GetProperty("availableItems").GetArrayLength());
        Assert.Equal("RevenuePivot", result.RootElement.GetProperty("connectedPivotTables")[0].GetString());
    }

    [Fact]
    public async Task Selection_UnknownCaptionPreservesServiceFailure()
    {
        var call = await fixture.CallToolAsync("slicer", new()
        {
            ["action"] = "set-slicer-selection",
            ["session_id"] = "session-1",
            ["slicer_name"] = "QuarterSlicer",
            ["selected_items"] = "[\"missing\"]"
        }, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "InvalidInput",
            ExceptionType = nameof(ArgumentException),
            ErrorMessage = "Slicer item 'missing' was not found."
        }, "slicer.set-slicer-selection",
            """{"slicerName":"QuarterSlicer","selectedItems":["missing"]}""");

        Assert.True(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("missing", result.RootElement.GetProperty("errorMessage").GetString());
    }

    [Fact]
    public async Task Selection_AmbiguousErrorPreservesCandidateJsonAndRetryName()
    {
        string[] candidates =
        [
            "[Calendar].[Quarter].&[2025]&[Q\"1\\North]]]",
            "[Calendar].[Quarter].&[2026]&[Q\"1\\South]]]"
        ];
        const string marker = "matching MDX unique names: ";
        string message = new ArgumentException(
            $"Slicer caption 'Q1' is ambiguous. Retry with one of these {marker}{JsonSerializer.Serialize(candidates)}",
            "requested").Message;
        var call = await fixture.CallToolAsync("slicer", new()
        {
            ["action"] = "set-slicer-selection",
            ["session_id"] = "session-1",
            ["slicer_name"] = "QuarterSlicer",
            ["selected_items"] = "[\"Q1\"]"
        }, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "InvalidInput",
            ExceptionType = nameof(ArgumentException),
            ErrorMessage = message
        }, "slicer.set-slicer-selection",
            """{"slicerName":"QuarterSlicer","selectedItems":["Q1"]}""");
        Assert.True(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", result.RootElement.GetProperty("errorCategory").GetString());
        string returnedMessage = result.RootElement.GetProperty("errorMessage").GetString()!;
        Assert.Equal(message, returnedMessage);
        var returnedCandidates = JsonSerializer.Deserialize<string[]>(returnedMessage[
            (returnedMessage.IndexOf(marker, StringComparison.Ordinal) + marker.Length)..(returnedMessage.LastIndexOf(']') + 1)])!;
        Assert.Equal(candidates, returnedCandidates);

        var retry = await fixture.CallToolAsync("slicer", new()
        {
            ["action"] = "set-slicer-selection",
            ["session_id"] = "session-1",
            ["slicer_name"] = "QuarterSlicer",
            ["selected_items"] = JsonSerializer.Serialize(new[] { returnedCandidates[1] })
        }, RecordingToolTest.Success("""{"success":true,"selectedItems":["Q1"]}"""),
            "slicer.set-slicer-selection", JsonSerializer.Serialize(new
            {
                slicerName = "QuarterSlicer",
                selectedItems = new[] { candidates[1] }
            }));
        Assert.False(retry.Result.IsError);
        using var success = JsonDocument.Parse(retry.JsonResult);
        Assert.True(success.RootElement.GetProperty("success").GetBoolean());
    }
}

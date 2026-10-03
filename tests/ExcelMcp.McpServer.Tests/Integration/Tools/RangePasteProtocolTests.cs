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
public sealed class RangePasteProtocolTests(RecordingProgramTransportFixture fixture)
{
    public static TheoryData<string, bool, bool> PasteOptions
    {
        get
        {
            TheoryData<string, bool, bool> data = [];
            foreach (var kind in new[] { "all", "values", "formulas", "formats", "validation" })
                foreach (var transpose in new[] { false, true })
                    foreach (var skipBlanks in new[] { false, true })
                        data.Add(kind, transpose, skipBlanks);
            return data;
        }
    }

    [Theory]
    [MemberData(nameof(PasteOptions))]
    public async Task Copy_MapsNativePasteOptionsAndReturnsResolvedBounds(string pasteKind, bool transpose, bool skipBlanks)
    {
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            destinationAddress = transpose ? "$D$1:$F$2" : "$D$1:$E$3",
            pasteKind,
            transpose,
            skipBlanks
        });
        var expected = JsonSerializer.Serialize(new
        {
            sourceSheet = "Source",
            sourceRange = "A1:B3",
            targetSheet = "Target",
            targetRange = "D1",
            pasteKind,
            transpose,
            skipBlanks
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range", new()
        {
            ["action"] = "copy",
            ["session_id"] = "session-1",
            ["source_sheet"] = "Source",
            ["source_range"] = "A1:B3",
            ["target_sheet"] = "Target",
            ["target_range"] = "D1",
            ["paste_kind"] = pasteKind,
            ["transpose"] = transpose,
            ["skip_blanks"] = skipBlanks
        }, RecordingToolTest.Success(response), "range.copy", expected);
        Assert.False(call.Result.IsError);
        using var output = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(transpose ? "$D$1:$F$2" : "$D$1:$E$3",
            output.RootElement.GetProperty("destinationAddress").GetString());
    }

    [Theory]
    [InlineData("copy-values")]
    [InlineData("copy-formulas")]
    public async Task RemovedCopyAction_IsRejected(string action)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range", new()
        {
            ["action"] = action,
            ["session_id"] = "session-1"
        });
        Assert.True(result.IsError);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("unknown")]
    [InlineData("99")]
    public async Task Copy_MissingOrInvalidKindDoesNotDispatch(string? pasteKind)
    {
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = "copy",
            ["session_id"] = "session-1",
            ["source_sheet"] = "Source",
            ["source_range"] = "A1:B3",
            ["target_sheet"] = "Target",
            ["target_range"] = "D1"
        };
        if (pasteKind is not null)
            arguments["paste_kind"] = pasteKind;
        var result = await fixture.CallResultWithoutDispatchAsync("range", arguments);
        Assert.True(result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesRequiredKindAndContentPreservation()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range");
        Assert.Contains("Required paste_kind", tool.Description, StringComparison.Ordinal);
        Assert.Contains("Formats/validation preserve content", tool.Description, StringComparison.Ordinal);
        Assert.DoesNotContain("copy-values", tool.Description, StringComparison.Ordinal);
        Assert.DoesNotContain("copy-formulas", tool.Description, StringComparison.Ordinal);
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("paste_kind", out _));
        Assert.True(properties.TryGetProperty("transpose", out _));
        Assert.True(properties.TryGetProperty("skip_blanks", out _));
    }
}

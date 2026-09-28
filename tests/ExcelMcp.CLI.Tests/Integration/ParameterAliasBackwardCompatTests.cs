using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Real parser coverage for backward-compatible range parameter aliases.
/// Workbook behavior is owned by the corresponding Service range tests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ParameterAliasBackwardCompatTests
{
    [Fact]
    public async Task RangeSetValues_ShortAliases_MapToCanonicalRequest()
    {
        var (result, requests) = await InProcessCliHelper.RunRecordingAsync(
            "range set-values --session session-1 --sheet AliasSheet --range C7 --values \"[[\\\"BackwardCompat\\\"]]\"",
            new ServiceResponse
            {
                Success = true,
                Result = """{"success":true}"""
            });

        Assert.Equal(0, result.ExitCode);
        var request = Assert.Single(requests);
        Assert.Equal("range.set-values", request.Command);
        Assert.Equal("session-1", request.SessionId);
        AssertJsonEqual(
            """{"sheetName":"AliasSheet","rangeAddress":"C7","values":[["BackwardCompat"]]}""",
            request.Args);
        AssertJsonEqual("""{"success":true}""", result.Stdout);
    }

    [Fact]
    public async Task RangeGetValues_ShortAliases_MapToCanonicalRequestAndResult()
    {
        var (result, requests) = await InProcessCliHelper.RunRecordingAsync(
            "range get-values --session session-1 --sheet AliasSheet --range E9",
            new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"values":[["ReadTest"]]}"""
            });

        Assert.Equal(0, result.ExitCode);
        var request = Assert.Single(requests);
        Assert.Equal("range.get-values", request.Command);
        Assert.Equal("session-1", request.SessionId);
        AssertJsonEqual(
            """{"sheetName":"AliasSheet","rangeAddress":"E9"}""",
            request.Args);
        AssertJsonEqual(
            """{"success":true,"values":[["ReadTest"]]}""",
            result.Stdout);
    }

    private static void AssertJsonEqual(string expectedJson, string? actualJson)
    {
        Assert.NotNull(actualJson);
        using var expected = JsonDocument.Parse(expectedJson);
        using var actual = JsonDocument.Parse(actualJson);
        Assert.True(
            JsonElement.DeepEquals(expected.RootElement, actual.RootElement),
            $"Expected {expected.RootElement.GetRawText()}, " +
            $"but received {actual.RootElement.GetRawText()}.");
    }
}

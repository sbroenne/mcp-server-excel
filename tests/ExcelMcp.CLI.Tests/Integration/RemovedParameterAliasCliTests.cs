using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Real parser coverage: only the canonical --sheet-name and --range-address names are accepted.
/// Workbook behavior is owned by the corresponding Service range tests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RemovedParameterAliasCliTests
{
    [Theory]
    [InlineData("range get-values --session session-1 --sheet AliasSheet --range-address E9", "sheet")]
    [InlineData("range get-values --session session-1 --sheet-name AliasSheet --range E9", "range")]
    public async Task RemovedShortAlias_IsRejectedWithoutDispatch(string arguments, string removedOption)
    {
        var (result, requests) = await InProcessCliHelper.RunRecordingAsync(arguments);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Empty(requests);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal(
            $"Unknown option '{removedOption}'.",
            output.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public async Task RangeGetValues_CanonicalNames_MapToRequestAndResult()
    {
        var (result, requests) = await InProcessCliHelper.RunRecordingAsync(
            "range get-values --session session-1 --sheet-name AliasSheet --range-address E9",
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
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("range.get-values", output.RootElement.GetProperty("command").GetString());
        AssertJsonEqual("""{"values":[["ReadTest"]]}""", output.RootElement.GetProperty("data").GetRawText());
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

using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Real CLI process coverage for backward-compatible range parameter aliases.
/// Canonical generated parameter mapping is covered by generated contract tests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class ParameterAliasBackwardCompatTests(
    CliWorkbookSessionFixture fixture,
    ITestOutputHelper output)
    : IClassFixture<CliWorkbookSessionFixture>
{
    private readonly CliWorkbookSessionFixture _fixture = fixture;
    private readonly ITestOutputHelper _output = output;

    [Fact]
    public async Task RangeSetValues_ShortAliases_WorksWithoutError()
    {
        var sessionId = _fixture.SessionId;
        var sheetName = await CreateSheetAsync(sessionId);

        try
        {
            var result = await CliProcessHelper.RunAsync(
                $"range set-values --session {sessionId} --sheet {sheetName} --range C7 --values \"[[\\\"BackwardCompat\\\"]]\"");

            _output.WriteLine($"Exit code: {result.ExitCode}");
            _output.WriteLine($"Stdout: {result.Stdout}");
            _output.WriteLine($"Stderr: {result.Stderr}");
            Assert.Equal(0, result.ExitCode);

            using var json = JsonDocument.Parse(result.Stdout);
            Assert.True(
                json.RootElement.GetProperty("success").GetBoolean(),
                "CLI should accept --sheet and --range aliases");

            var (read, readJson) = await CliProcessHelper.RunJsonAsync(
                ["range", "get-values", "--session", sessionId,
                 "--sheet-name", sheetName, "--range-address", "C7"]);
            Assert.Equal(0, read.ExitCode);
            Assert.True(readJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(
                "BackwardCompat",
                readJson.RootElement.GetProperty("values")[0][0].GetString());
        }
        finally
        {
            await DeleteSheetAsync(sessionId, sheetName);
        }
    }

    [Fact]
    public async Task RangeGetValues_ShortAliases_WorksWithoutError()
    {
        var sessionId = _fixture.SessionId;
        var sheetName = await CreateSheetAsync(sessionId);

        try
        {
            var (setup, setupJson) = await CliProcessHelper.RunJsonAsync(
                ["range", "set-values", "--session", sessionId,
                 "--sheet-name", sheetName, "--range-address", "E9",
                 "--values", "[[\"ReadTest\"]]"]);
            Assert.Equal(0, setup.ExitCode);
            Assert.True(setupJson.RootElement.GetProperty("success").GetBoolean());

            var result = await CliProcessHelper.RunAsync(
                $"range get-values --session {sessionId} --sheet {sheetName} --range E9");

            _output.WriteLine($"Exit code: {result.ExitCode}");
            _output.WriteLine($"Stdout: {result.Stdout}");
            Assert.Equal(0, result.ExitCode);

            using var json = JsonDocument.Parse(result.Stdout);
            Assert.True(
                json.RootElement.GetProperty("success").GetBoolean(),
                "CLI should accept --sheet and --range aliases for get-values");
            Assert.Equal(
                "ReadTest",
                json.RootElement.GetProperty("values")[0][0].GetString());
        }
        finally
        {
            await DeleteSheetAsync(sessionId, sheetName);
        }
    }

    private static async Task<string> CreateSheetAsync(string sessionId)
    {
        var sheetName = $"Alias_{Guid.NewGuid():N}"[..21];
        var result = await CliProcessHelper.RunAsync(
            ["sheet", "create", "--session", sessionId, "--sheet-name", sheetName]);
        Assert.True(
            result.ExitCode == 0,
            $"CLI sheet creation failed: {result.Stdout}{result.Stderr}");
        return sheetName;
    }

    private static async Task DeleteSheetAsync(
        string sessionId,
        string sheetName)
    {
        var result = await CliProcessHelper.RunAsync(
            ["sheet", "delete", "--session", sessionId, "--sheet-name", sheetName]);
        Assert.True(
            result.ExitCode == 0,
            $"CLI sheet cleanup failed: {result.Stdout}{result.Stderr}");
    }
}

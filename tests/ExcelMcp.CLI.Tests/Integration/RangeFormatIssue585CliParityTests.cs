using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Real CLI argument, output, and exit-code coverage for range formatting.
/// Formatting behavior and persistence are covered by Service tests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class RangeFormatIssue585CliParityTests(
    ITestOutputHelper output)
    : IAsyncLifetime
{
    private readonly CliWorkbookSessionFixture _fixture = new();
    private readonly ITestOutputHelper _output = output;

    public Task InitializeAsync() => _fixture.InitializeAsync();
    public Task DisposeAsync() => _fixture.DisposeAsync();

    [Fact]
    public async Task FormatRanges_NumberFormat_RoundTripsInvariantCodeViaCli()
    {
        var sessionId = _fixture.SessionId;
        var sheetName = await CreateSheetAsync(sessionId);

        var seed = await CliProcessHelper.RunAsync(
            ["range", "set-values", "--session", sessionId, "--sheet-name", sheetName,
                    "--range-address", "A1:A2", "--values", "[[0.25],[0.5]]"]);
        Assert.True(seed.ExitCode == 0, seed.Stdout + seed.Stderr);
        var (formatted, formatJson) = await CliProcessHelper.RunJsonAsync(
            ["rangeformat", "format", "--session", sessionId,
                 "--sheet-name", sheetName, "--range-addresses", "A1:A2",
                 "--format-options", """{"numberFormat":"0.00%"}"""],
            timeoutMs: 60000);
        Assert.True(
            formatted.ExitCode == 0,
            formatted.Stdout + formatted.Stderr);
        Assert.True(formatJson.RootElement.GetProperty("success").GetBoolean());

        var (read, readJson) = await CliProcessHelper.RunJsonAsync(
            ["range", "get-number-formats", "--session", sessionId,
                 "--sheet-name", sheetName, "--range-address", "A1:A2"],
            timeoutMs: 60000);
        Assert.Equal(0, read.ExitCode);
        Assert.True(readJson.RootElement.GetProperty("success").GetBoolean());
        var formats = readJson.RootElement.GetProperty("formats");
        Assert.Equal(2, formats.GetArrayLength());
        foreach (var row in formats.EnumerateArray())
        {
            Assert.Equal(1, row.GetArrayLength());
            Assert.Equal("0.00%", row[0].GetString());
        }
        var (valuesResult, valuesDocument) = await CliProcessHelper.RunJsonAsync(
            ["range", "get-values", "--session", sessionId, "--sheet-name", sheetName,
                    "--range-address", "A1:A2"]);
        using (valuesDocument)
        {
            Assert.Equal(0, valuesResult.ExitCode);
            Assert.True(valuesDocument.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(0.25, valuesDocument.RootElement.GetProperty("values")[0][0].GetDouble());
            Assert.Equal(0.5, valuesDocument.RootElement.GetProperty("values")[1][0].GetDouble());
        }
    }

    [Fact]
    public async Task FormatRange_InvalidColor_ReturnsTransparentFailureEnvelopeViaCli()
    {
        var sessionId = _fixture.SessionId;
        var sheetName = await CreateSheetAsync(sessionId);

        var seeded = await CliProcessHelper.RunAsync(
            ["range", "set-values", "--session", sessionId, "--sheet-name", sheetName,
                    "--range-address", "A1:J1", "--values", "[[1,2,3,4,5,6,7,8,9,10]]"]);
        Assert.True(seeded.ExitCode == 0, seeded.Stdout + seeded.Stderr);
        var formatted = await CliProcessHelper.RunAsync(
            ["rangeformat", "format", "--session", sessionId, "--sheet-name", sheetName,
                    "--range-addresses", "A1:J1", "--format-options", """{"fillColor":"#123456","numberFormat":"0.00%"}"""]);
        Assert.True(formatted.ExitCode == 0, formatted.Stdout + formatted.Stderr);
        var before = await ReadStateAsync(sessionId, sheetName);
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["rangeformat", "format", "--session", sessionId,
                 "--sheet-name", sheetName, "--range-addresses", "A1:J1",
                 "--format-options", """{"fillColor":"not-a-color"}"""],
            timeoutMs: 60000);

        _output.WriteLine($"CLI stdout: {result.Stdout}");
        _output.WriteLine($"CLI stderr: {result.Stderr}");
        Assert.Equal(1, result.ExitCode);

        var root = json.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(
            "ArgumentException",
            root.GetProperty("exceptionType").GetString());
        Assert.Equal(
            "InvalidInput",
            root.GetProperty("errorCategory").GetString());
        Assert.Equal(
            "rangeformat.format",
            root.GetProperty("command").GetString());
        Assert.Equal(
            sessionId,
            root.GetProperty("sessionId").GetString());
        Assert.Contains(
            "Invalid color format: not-a-color",
            root.GetProperty("errorMessage").GetString(),
            StringComparison.Ordinal);
        Assert.Equal(before, await ReadStateAsync(sessionId, sheetName));
    }

    private static async Task<string> CreateSheetAsync(string sessionId)
    {
        var sheetName = $"CliFmt_{Guid.NewGuid():N}"[..22];
        var result = await CliProcessHelper.RunAsync(
            ["sheet", "create", "--session", sessionId, "--sheet-name", sheetName],
            timeoutMs: 60000);
        Assert.True(
            result.ExitCode == 0,
            $"CLI sheet creation failed: {result.Stdout}{result.Stderr}");
        return sheetName;
    }

    private static async Task<string> ReadStateAsync(string sessionId, string sheetName)
    {
        var (values, valuesDocument) = await CliProcessHelper.RunJsonAsync(
            ["range", "get-values", "--session", sessionId, "--sheet-name", sheetName,
                "--range-address", "A1:J1"]);
        using var retainedValues = valuesDocument;
        Assert.Equal(0, values.ExitCode);
        Assert.True(retainedValues.RootElement.GetProperty("success").GetBoolean());
        var (formats, formatDocument) = await CliProcessHelper.RunJsonAsync(
            ["rangeformat", "get-format", "--session", sessionId, "--sheet-name", sheetName,
                "--range-address", "A1:J1"]);
        using var retainedFormats = formatDocument;
        Assert.True(formats.ExitCode == 0, formats.Stdout + formats.Stderr);
        Assert.True(retainedFormats.RootElement.GetProperty("success").GetBoolean());
        return retainedValues.RootElement.GetRawText() + retainedFormats.RootElement.GetRawText();
    }
}

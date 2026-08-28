using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class RangeTypedIsoValueCliTests : IDisposable
{
    private readonly string _testFile = Path.Combine(
        Path.GetTempPath(),
        $"RangeTypedIsoValueCli_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public async Task RangeSetValues_TypedIsoDate_RoundTripsThroughCli()
    {
        var (createResult, createJson) = await CliProcessHelper.RunJsonAsync(
            ["session", "create", _testFile],
            timeoutMs: 60000,
            diagnosticLabel: "create typed ISO workbook");
        Assert.Equal(0, createResult.ExitCode);
        var sessionId = createJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        var activeSessionId = sessionId;

        try
        {
            const string valuesJson =
                """[[{"type":"datetime-offset","value":"2026-08-27T03:15:00-04:00","numberFormat":"@"},"2026-08-27"]]""";
            var (setResult, setJson) = await CliProcessHelper.RunJsonAsync(
                ["range", "set-values", "--session", sessionId!, "--sheet-name", "Sheet1", "--range-address", "A1:B1", "--values", valuesJson],
                timeoutMs: 60000,
                diagnosticLabel: "set typed ISO value");
            Assert.Equal(0, setResult.ExitCode);
            Assert.True(setJson.RootElement.GetProperty("success").GetBoolean(), setResult.Stdout);

            var (getResult, getJson) = await CliProcessHelper.RunJsonAsync(
                ["range", "get-values", "--session", sessionId!, "--sheet-name", "Sheet1", "--range-address", "A1:B1"],
                timeoutMs: 60000,
                diagnosticLabel: "get typed ISO value");
            Assert.Equal(0, getResult.ExitCode);
            Assert.Equal(
                new DateTime(2026, 8, 27, 7, 15, 0).ToOADate(),
                getJson.RootElement.GetProperty("values")[0][0].GetDouble());
            Assert.Equal("2026-08-27", getJson.RootElement.GetProperty("values")[0][1].GetString());

            var (formatResult, formatJson) = await CliProcessHelper.RunJsonAsync(
                ["range", "get-number-formats", "--session", sessionId!, "--sheet-name", "Sheet1", "--range-address", "A1:B1"],
                timeoutMs: 60000,
                diagnosticLabel: "get typed ISO formats");
            Assert.Equal(0, formatResult.ExitCode);
            Assert.Equal(
                "@",
                formatJson.RootElement.GetProperty("formats")[0][0].GetString(),
                StringComparer.OrdinalIgnoreCase);

            var closeResult = await CliProcessHelper.RunAsync(
                ["session", "close", "--session", activeSessionId!, "--save", "true"],
                timeoutMs: 60000,
                diagnosticLabel: "save typed ISO workbook");
            Assert.Equal(0, closeResult.ExitCode);
            activeSessionId = null;

            var (openResult, openJson) = await CliProcessHelper.RunJsonAsync(
                ["session", "open", _testFile],
                timeoutMs: 60000,
                diagnosticLabel: "reopen typed ISO workbook");
            Assert.Equal(0, openResult.ExitCode);
            activeSessionId = openJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(activeSessionId));

            var (persistedResult, persistedJson) = await CliProcessHelper.RunJsonAsync(
                ["range", "get-values", "--session", activeSessionId!, "--sheet-name", "Sheet1", "--range-address", "A1:B1"],
                timeoutMs: 60000,
                diagnosticLabel: "get persisted typed ISO value");
            Assert.Equal(0, persistedResult.ExitCode);
            Assert.Equal(
                new DateTime(2026, 8, 27, 7, 15, 0).ToOADate(),
                persistedJson.RootElement.GetProperty("values")[0][0].GetDouble());
            Assert.Equal("2026-08-27", persistedJson.RootElement.GetProperty("values")[0][1].GetString());
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(activeSessionId))
            {
                await CliProcessHelper.RunAsync(
                    ["session", "close", "--session", activeSessionId, "--save", "false"],
                    timeoutMs: 60000,
                    diagnosticLabel: "close typed ISO workbook");
            }
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFile))
        {
            File.Delete(_testFile);
        }

        GC.SuppressFinalize(this);
    }
}

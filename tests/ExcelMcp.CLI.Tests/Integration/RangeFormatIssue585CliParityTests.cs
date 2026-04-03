using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class RangeFormatIssue585CliParityTests : IDisposable
{
    private const int IssueFillColor = 7949855;
    private const int WhiteFontColor = 16777215;

    private readonly ITestOutputHelper _output;
    private readonly string _testFile;

    public RangeFormatIssue585CliParityTests(ITestOutputHelper output)
    {
        _output = output;
        _testFile = Path.Combine(Path.GetTempPath(), $"RangeFormatIssue585Cli_{Guid.NewGuid():N}.xlsx");
    }

    [Fact]
    public async Task RangeFormat_FormatRange_Issue585Payload_SucceedsViaCli()
    {
        var (openResult, openJson) = await CliProcessHelper.RunJsonAsync(
            $"session create \"{_testFile}\"",
            timeoutMs: 60000,
            diagnosticLabel: "session create for rangeformat parity");
        Assert.Equal(0, openResult.ExitCode);

        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        try
        {
            var createSheetResult = await CliProcessHelper.RunAsync(
                $"sheet create --session {sessionId} --sheet-name \"Toutes les transactions\"",
                timeoutMs: 60000,
                diagnosticLabel: "sheet create for rangeformat parity");
            Assert.Equal(0, createSheetResult.ExitCode);

            var formatResult = await CliProcessHelper.RunAsync(
                $"rangeformat format-range --session {sessionId} --sheet-name \"Toutes les transactions\" --range-address A1:J1 --bold true --fill-color \"#1F4E79\" --font-color \"#FFFFFF\"",
                timeoutMs: 60000,
                diagnosticLabel: "rangeformat format-range issue585 payload");

            _output.WriteLine($"CLI stdout: {formatResult.Stdout}");
            _output.WriteLine($"CLI stderr: {formatResult.Stderr}");

            Assert.Equal(0, formatResult.ExitCode);

            using var formatJson = JsonDocument.Parse(formatResult.Stdout);
            Assert.True(formatJson.RootElement.GetProperty("success").GetBoolean(), "CLI rangeformat should succeed for the issue #585 payload.");
        }
        finally
        {
            await CliProcessHelper.RunAsync(
                $"session close --session {sessionId} --save true",
                timeoutMs: 60000,
                diagnosticLabel: "session close for rangeformat parity");
        }

        using var batch = ExcelSession.BeginBatch(_testFile);
        var a1 = ReadCellFormatting(batch, "Toutes les transactions", "A1");
        var j1 = ReadCellFormatting(batch, "Toutes les transactions", "J1");

        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), a1);
        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), j1);
    }

    [Fact]
    public async Task RangeFormat_FormatRange_InvalidColor_ReturnsTransparentFailureEnvelopeViaCli()
    {
        var (openResult, openJson) = await CliProcessHelper.RunJsonAsync(
            $"session create \"{_testFile}\"",
            timeoutMs: 60000,
            diagnosticLabel: "session create for rangeformat diagnostics");
        Assert.Equal(0, openResult.ExitCode);

        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        try
        {
            var createSheetResult = await CliProcessHelper.RunAsync(
                $"sheet create --session {sessionId} --sheet-name \"Toutes les transactions\"",
                timeoutMs: 60000,
                diagnosticLabel: "sheet create for rangeformat diagnostics");
            Assert.Equal(0, createSheetResult.ExitCode);

            var (formatResult, formatJson) = await CliProcessHelper.RunJsonAsync(
                $"rangeformat format-range --session {sessionId} --sheet-name \"Toutes les transactions\" --range-address A1:J1 --fill-color not-a-color",
                timeoutMs: 60000,
                diagnosticLabel: "rangeformat invalid-color diagnostics");

            _output.WriteLine($"CLI stdout: {formatResult.Stdout}");
            _output.WriteLine($"CLI stderr: {formatResult.Stderr}");

            Assert.Equal(1, formatResult.ExitCode);
            AssertFailureEnvelope(
                formatJson.RootElement,
                "CLI rangeformat invalid-color",
                expectedExceptionType: "ArgumentException",
                expectedErrorCategory: "InvalidInput",
                expectedCommand: "rangeformat.format-range",
                expectedSessionId: sessionId);

            var errorMessage = formatJson.RootElement.GetProperty("errorMessage").GetString();
            Assert.Contains("Invalid color format: not-a-color", errorMessage, StringComparison.Ordinal);
        }
        finally
        {
            await CliProcessHelper.RunAsync(
                $"session close --session {sessionId} --save false",
                timeoutMs: 60000,
                diagnosticLabel: "session close for rangeformat diagnostics");
        }
    }

    [Fact]
    public async Task RangeFormat_FormatRange_Issue585Payload_OnMissingSheet_ReturnsTransparentFailure()
    {
        var (openResult, openJson) = await CliProcessHelper.RunJsonAsync(
            $"session create \"{_testFile}\"",
            timeoutMs: 60000,
            diagnosticLabel: "session create for rangeformat failure parity");
        Assert.Equal(0, openResult.ExitCode);

        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        try
        {
            var formatResult = await CliProcessHelper.RunAsync(
                $"rangeformat format-range --session {sessionId} --sheet-name \"Toutes les transactions\" --range-address A1:J1 --bold true --fill-color \"#1F4E79\" --font-color \"#FFFFFF\"",
                timeoutMs: 60000,
                diagnosticLabel: "rangeformat format-range issue585 failure payload");

            _output.WriteLine($"CLI stdout: {formatResult.Stdout}");
            _output.WriteLine($"CLI stderr: {formatResult.Stderr}");

            Assert.Equal(1, formatResult.ExitCode);

            using var formatJson = JsonDocument.Parse(formatResult.Stdout);
            AssertFailureEnvelope(
                formatJson.RootElement,
                "CLI rangeformat missing-sheet",
                expectedExceptionType: "COMException",
                expectedErrorCategory: "ComInterop",
                expectedCommand: "rangeformat.format-range",
                expectedSessionId: sessionId);

            Assert.True(formatJson.RootElement.TryGetProperty("hresult", out var hresult));
            Assert.False(string.IsNullOrWhiteSpace(hresult.GetString()));

            var errorMessage = formatJson.RootElement.GetProperty("errorMessage").GetString();
            Assert.NotNull(errorMessage);
            Assert.Contains("COMException", errorMessage, StringComparison.Ordinal);
            Assert.Contains("Invalid index", errorMessage, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            await CliProcessHelper.RunAsync(
                $"session close --session {sessionId} --save false",
                timeoutMs: 60000,
                diagnosticLabel: "session close for rangeformat failure parity");
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFile))
        {
            try
            {
                File.Delete(_testFile);
            }
            catch
            {
            }
        }

        GC.SuppressFinalize(this);
    }

    private static CellFormattingState ReadCellFormatting(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            dynamic? interior = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                font = range.Font;
                interior = range.Interior;

                return new CellFormattingState(
                    Bold: Convert.ToBoolean(font.Bold),
                    FillColor: interior.Color == null ? null : Convert.ToInt32(interior.Color),
                    FontColor: font.Color == null ? null : Convert.ToInt32(font.Color));
            }
            finally
            {
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private sealed record CellFormattingState(bool Bold, int? FillColor, int? FontColor);

    private static void AssertFailureEnvelope(
        JsonElement root,
        string operationName,
        string expectedExceptionType,
        string? expectedErrorCategory = null,
        string? expectedCommand = null,
        string? expectedSessionId = null)
    {
        Assert.False(root.GetProperty("success").GetBoolean(), $"{operationName} unexpectedly succeeded.");
        Assert.True(root.GetProperty("isError").GetBoolean(), $"{operationName} should return isError=true.");
        Assert.Equal(expectedExceptionType, root.GetProperty("exceptionType").GetString());

        var error = root.GetProperty("error").GetString();
        var errorMessage = root.GetProperty("errorMessage").GetString();

        Assert.False(string.IsNullOrWhiteSpace(errorMessage), $"{operationName} should return errorMessage.");
        Assert.Equal(errorMessage, error);
        Assert.False(root.TryGetProperty("innerError", out _), $"{operationName} should not include innerError for these failures.");
        AssertOptionalStringProperty(root, "errorCategory", expectedErrorCategory, operationName);
        AssertOptionalStringProperty(root, "command", expectedCommand, operationName);
        AssertOptionalStringProperty(root, "sessionId", expectedSessionId, operationName);
    }

    private static void AssertOptionalStringProperty(
        JsonElement root,
        string propertyName,
        string? expectedValue,
        string operationName)
    {
        if (expectedValue == null)
        {
            Assert.False(root.TryGetProperty(propertyName, out _), $"{operationName} should not include {propertyName}.");
            return;
        }

        Assert.True(root.TryGetProperty(propertyName, out var property), $"{operationName} should include {propertyName}.");
        Assert.Equal(expectedValue, property.GetString());
    }
}

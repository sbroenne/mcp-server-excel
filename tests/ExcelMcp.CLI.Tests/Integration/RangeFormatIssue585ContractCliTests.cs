using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// CLI arguments and failure output for range formatting (issue 585). Number
/// formats and invalid colors in Excel are covered by
/// PersistentServiceRangeSpecializedTests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFormatIssue585ContractCliTests
{
    [Fact]
    public async Task Format_SendsNumberFormatCodeUnchanged()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["rangeformat", "format", "--session", "session-1", "--sheet-name", "Sheet1",
                "--range-addresses", "A1:A2", "--format-options", """{"numberFormat":"0.00%"}"""],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeformat.format", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1:A2", Assert.Single(args.RootElement.GetProperty("rangeAddresses").EnumerateArray()).GetString());
        Assert.Equal("0.00%", args.RootElement.GetProperty("formatOptions").GetProperty("numberFormat").GetString());
    }

    [Fact]
    public async Task Format_ServiceFailure_ReturnsFullFailureEnvelopeAndExitCodeOne()
    {
        var result = await InProcessCliHelper.RunAsync(
            ["rangeformat", "format", "--session", "session-1", "--sheet-name", "Sheet1",
                "--range-addresses", "A1:J1", "--format-options", """{"fillColor":"not-a-color"}"""],
            _ => new ServiceResponse
            {
                Success = false,
                Command = "rangeformat.format",
                SessionId = "session-1",
                ErrorMessage = "Invalid color format: not-a-color",
                ErrorCategory = "InvalidInput",
                ExceptionType = nameof(ArgumentException)
            });

        Assert.Equal(1, result.ExitCode);
        using var document = JsonDocument.Parse(result.Stdout);
        var root = document.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(nameof(ArgumentException), root.GetProperty("error").GetProperty("exceptionType").GetString());
        Assert.Equal("InvalidInput", root.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal("rangeformat.format", root.GetProperty("command").GetString());
        Assert.Equal("session-1", root.GetProperty("error").GetProperty("sessionId").GetString());
        Assert.Contains("Invalid color format: not-a-color", root.GetProperty("error").GetProperty("message").GetString(), StringComparison.Ordinal);
    }
}

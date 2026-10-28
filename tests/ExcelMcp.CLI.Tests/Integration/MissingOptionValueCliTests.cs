using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Session")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class MissingOptionValueCliTests
{
    // Spectre retries a failed parse by inserting its internal default-command name,
    // which can otherwise become the value of an option whose value was omitted.
    [Theory]
    [InlineData("--session", "sheet", "list", "--session")]
    [InlineData("-s", "sheet", "list", "-s")]
    [InlineData("--session", "range", "get-values", "--session", "--sheet-name", "Sheet1", "--range-address", "A1")]
    [InlineData("--sheet-name", "range", "get-values", "--session", "session-1", "--range-address", "A1", "--sheet-name")]
    [InlineData("--session", "session", "close", "--session")]
    [InlineData("--workbook-id", "session", "attach", "--workbook-id")]
    public async Task OptionWithoutValue_IsRejectedWithoutDispatch(string option, params string[] arguments)
    {
        var dispatched = 0;
        var result = await InProcessCliHelper.RunAsync(arguments, _ =>
        {
            dispatched++;
            return new ServiceResponse { Success = true };
        });

        Assert.NotEqual(0, result.ExitCode);
        Assert.Equal(0, dispatched);
        using var output = JsonDocument.Parse(result.Stdout);
        var error = output.RootElement.GetProperty("error");
        Assert.Equal("InvalidInput", error.GetProperty("category").GetString());
        Assert.Equal($"Option '{option}' requires a value.", error.GetProperty("message").GetString());
        Assert.DoesNotContain("__default_command", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("session", "open")]
    [InlineData("session", "create")]
    public async Task MissingFilePath_IsRejectedWithoutDispatch(params string[] arguments)
    {
        var dispatched = 0;
        var result = await InProcessCliHelper.RunAsync(arguments, _ =>
        {
            dispatched++;
            return new ServiceResponse { Success = true };
        });

        Assert.NotEqual(0, result.ExitCode);
        Assert.Equal(0, dispatched);
        Assert.DoesNotContain("__default_command", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public async Task OptionValue_TypedByUser_IsForwardedUnchanged()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(["sheet", "list", "--session", "__default_command"], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, worksheets = Array.Empty<object>() })
            };
        });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("sheet.list", captured.Command);
        Assert.Equal("__default_command", captured.SessionId);
    }
}

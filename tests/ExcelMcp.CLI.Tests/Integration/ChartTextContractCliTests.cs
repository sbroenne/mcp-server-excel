using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ChartDepth")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartTextContractCliTests
{
    [Fact]
    public async Task SetTextFormat_PreservesExplicitFalseAndFontSettings()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "chartconfig", "set-text-format", "--session", "session-1", "--chart-name", "Sales",
                "--element", "Title", "--font-name", "Arial", "--font-size", "14.5", "--bold", "false",
                "--italic", "true", "--color", "#2468AC"], request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true,"formats":[]}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("chartconfig.set-text-format", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Arial", args.RootElement.GetProperty("fontName").GetString());
        Assert.Equal(14.5, args.RootElement.GetProperty("fontSize").GetDouble());
        Assert.False(args.RootElement.GetProperty("bold").GetBoolean());
        Assert.True(args.RootElement.GetProperty("italic").GetBoolean());
        Assert.Equal("#2468AC", args.RootElement.GetProperty("color").GetString());
    }

    [Fact]
    public async Task ReadTextFormat_PreservesSecondaryAxisAndNativeResults()
    {
        const string response = """{"success":true,"formats":[{"axis":"ValueSecondary","fontName":"Arial","fontSize":14.5,"bold":false,"italic":true,"color":"#2468AC","mixedFields":[]}]}""";
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "chartconfig", "get-text-format", "--session", "session-1", "--chart-name", "Sales",
                "--element", "ValueAxis", "--axis", "ValueSecondary"], request =>
            {
                Assert.Equal("chartconfig.get-text-format", request.Command);
                Assert.Equal("session-1", request.SessionId);
                using var args = JsonDocument.Parse(request.Args!);
                Assert.Equal("ValueAxis", args.RootElement.GetProperty("element").GetString());
                Assert.Equal("ValueSecondary", args.RootElement.GetProperty("axis").GetString());
                Assert.False(args.RootElement.TryGetProperty("fontSize", out _));
                return new ServiceResponse { Success = true, Result = response };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        using var output = JsonDocument.Parse(result.Stdout);
        using var expected = JsonDocument.Parse(response);
        Assert.True(JsonElement.DeepEquals(expected.RootElement.GetProperty("formats"), output.RootElement.GetProperty("data").GetProperty("formats")));
    }

    [Fact]
    public async Task TextFormat_ForwardsFailureWithoutSuccessData()
    {
        var result = await InProcessCliHelper.RunAsync(
            ["chartconfig", "set-text-format", "--session", "session-1", "--chart-name", "Sales",
                "--element", "Title", "--font-size", "14"], request => new ServiceResponse
                {
                    Success = false,
                    Command = request.Command,
                    SessionId = request.SessionId,
                    ExceptionType = "InvalidOperationException",
                    ErrorMessage = "The chart has no title."
                });
        Assert.Equal(1, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.False(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("The chart has no title.", output.RootElement.GetProperty("error").GetProperty("message").GetString());
        Assert.False(output.RootElement.TryGetProperty("data", out _));
    }
}

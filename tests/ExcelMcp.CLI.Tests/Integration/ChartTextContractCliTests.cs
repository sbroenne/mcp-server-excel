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
    [Theory]
    [InlineData("left", "Left")]
    [InlineData("center", "Center")]
    [InlineData("right", "Right")]
    public async Task SetTitle_ForwardsPositionAndAlignment(string alignment, string expectedAlignment)
    {
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "chartconfig", "set-title", "--session", "session-1", "--chart-name", "Sales",
                "--title", "Revenue", "--left", "16", "--top", "12", "--alignment", alignment], request =>
            {
                Assert.Equal("chartconfig.set-title", request.Command);
                Assert.Equal("session-1", request.SessionId);
                using var args = JsonDocument.Parse(request.Args!);
                Assert.Equal("Sales", args.RootElement.GetProperty("chartName").GetString());
                Assert.Equal("Revenue", args.RootElement.GetProperty("title").GetString());
                Assert.Equal(16, args.RootElement.GetProperty("left").GetDouble());
                Assert.Equal(12, args.RootElement.GetProperty("top").GetDouble());
                Assert.Equal(expectedAlignment, args.RootElement.GetProperty("alignment").GetString());
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LineFormatting_ForwardsVisibilityAndSecondaryAxis(bool area)
    {
        var action = area ? "set-area-format" : "set-axis-format";
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "chartconfig", action, "--session", "session-1", "--chart-name", "Sales",
                area ? "--area" : "--axis", area ? "Plot" : "ValueSecondary",
                "--line-color", "#D0D0D0", "--line-weight", "1.5", "--line-visible", "false"], request =>
            {
                Assert.Equal($"chartconfig.{action}", request.Command);
                using var args = JsonDocument.Parse(request.Args!);
                Assert.Equal(area ? "Plot" : "ValueSecondary", args.RootElement.GetProperty(area ? "area" : "axis").GetString());
                Assert.Equal("#D0D0D0", args.RootElement.GetProperty("lineColor").GetString());
                Assert.Equal(1.5, args.RootElement.GetProperty("lineWeight").GetDouble());
                Assert.False(args.RootElement.GetProperty("lineVisible").GetBoolean());
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
    }

    [Fact]
    public async Task TitleAndLineFormatting_OmittedSettingsStayOmitted()
    {
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "chartconfig", "set-title", "--session", "session-1", "--chart-name", "Sales",
                "--title", "Revenue"], request =>
            {
                using var args = JsonDocument.Parse(request.Args!);
                Assert.False(args.RootElement.TryGetProperty("left", out _));
                Assert.False(args.RootElement.TryGetProperty("top", out _));
                Assert.False(args.RootElement.TryGetProperty("alignment", out _));
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        result = await InProcessCliHelper.RunAsync(
            ["-q", "chartconfig", "set-axis-format", "--session", "session-1", "--chart-name", "Sales",
                "--axis", "Value", "--line-color", "#D0D0D0"], request =>
            {
                using var args = JsonDocument.Parse(request.Args!);
                Assert.False(args.RootElement.TryGetProperty("lineWeight", out _));
                Assert.False(args.RootElement.TryGetProperty("lineVisible", out _));
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
    }

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

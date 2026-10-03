using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CalculationMode")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CalculationSettingsContractCliTests
{
    [Fact]
    public async Task Settings_MapsPartialIterationChange()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "calculationmode", "set-settings", "--session", "session-1",
            "--iteration-enabled", "false", "--maximum-iterations", "37", "--maximum-change", "0.0002"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("calculation.set-settings", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.False(args.RootElement.GetProperty("iterationEnabled").GetBoolean());
        Assert.Equal(37, args.RootElement.GetProperty("maximumIterations").GetInt32());
        Assert.Equal(0.0002, args.RootElement.GetProperty("maximumChange").GetDouble());
        Assert.False(args.RootElement.TryGetProperty("mode", out _));
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("full")]
    [InlineData("rebuild")]
    public async Task Calculation_MapsExplicitNativeStrength(string kind)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "calculationmode", "calculate", "--session", "session-1",
            "--scope", "application", "--kind", kind
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("application", args.RootElement.GetProperty("scope").GetString(), ignoreCase: true);
        Assert.Equal(kind, args.RootElement.GetProperty("kind").GetString(), ignoreCase: true);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Precision_MapsExplicitLossPermission(bool enabled)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "calculationmode", "set-precision", "--session", "session-1",
            "--precision-as-displayed", enabled ? "true" : "false",
            "--allow-precision-loss", enabled ? "true" : "false"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(enabled, args.RootElement.GetProperty("precisionAsDisplayed").GetBoolean());
        Assert.Equal(enabled, args.RootElement.GetProperty("allowPrecisionLoss").GetBoolean());
    }

    [Theory]
    [InlineData("sheet", null, "sheetName")]
    [InlineData("range", null, "sheetName")]
    [InlineData("range", "Sheet1", "rangeAddress")]
    public async Task Calculation_MissingTargetForwardsInvalidInputAndNonzeroExit(
        string scope, string? sheetName, string parameter)
    {
        List<string> arguments =
        [
            "--quiet", "calculationmode", "calculate", "--session", "session-1", "--scope", scope
        ];
        if (sheetName is not null)
            arguments.AddRange(["--sheet", sheetName]);
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync([.. arguments], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = false,
                Command = "calculation.calculate",
                SessionId = "session-1",
                ErrorCategory = "InvalidInput",
                ExceptionType = nameof(ArgumentException),
                ErrorMessage = $"{parameter} is required for calculation."
            };
        });

        Assert.Equal(1, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal("calculation.calculate", captured.Command);
        using var json = JsonDocument.Parse(result.Stdout);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", json.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal(nameof(ArgumentException), json.RootElement.GetProperty("exceptionType").GetString());
        Assert.Contains(parameter, json.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("get-mode")]
    [InlineData("set-mode")]
    public async Task RemovedActions_DoNotDispatch(string action)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["calculationmode", action, "--session", "session-1"], request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.NotEqual(0, result.ExitCode);
        Assert.Null(captured);
    }
}

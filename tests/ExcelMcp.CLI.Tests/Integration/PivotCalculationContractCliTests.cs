using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotCalculation")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotCalculationContractCliTests
{
    [Theory]
    [InlineData("Named", "North")]
    [InlineData("Previous", null)]
    [InlineData("Next", null)]
    public async Task AdditionalCalculation_MapsExplicitBaseSettings(string kind, string? itemName)
    {
        ServiceRequest? captured = null;
        List<string> arguments =
        [
            "pivottablefield", "set-field-calculation", "--session", "session-1",
            "--pivot-table-name", "SalesPivot", "--field-name", "Total Sales",
            "--calculation", "DifferenceFrom", "--base-field-name", "Region",
            "--base-item-kind", kind
        ];
        if (itemName is not null)
            arguments.AddRange(["--base-item-name", itemName]);
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("pivottablefield.set-field-calculation", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Total Sales", args.RootElement.GetProperty("fieldName").GetString());
        Assert.Equal("DifferenceFrom", args.RootElement.GetProperty("calculation").GetString());
        Assert.Equal("Region", args.RootElement.GetProperty("baseFieldName").GetString());
        Assert.Equal(kind, args.RootElement.GetProperty("baseItemKind").GetString());
        if (itemName is not null)
            Assert.Equal(itemName, args.RootElement.GetProperty("baseItemName").GetString());
        else
            Assert.False(args.RootElement.TryGetProperty("baseItemName", out _));
    }

    [Fact]
    public async Task NormalReset_DoesNotInventBaseSettings()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "pivottablefield", "set-field-calculation", "--session", "session-1",
            "--pivot-table-name", "SalesPivot", "--field-name", "Average Sales", "--calculation", "Normal"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Normal", args.RootElement.GetProperty("calculation").GetString());
        Assert.False(args.RootElement.TryGetProperty("baseFieldName", out _));
        Assert.False(args.RootElement.TryGetProperty("baseItemKind", out _));
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData(null)]
    public async Task MissingOrUnknownCalculation_DoesNotDispatch(string? calculation)
    {
        bool dispatched = false;
        List<string> arguments =
        [
            "pivottablefield", "set-field-calculation", "--session", "session-1",
            "--pivot-table-name", "SalesPivot", "--field-name", "Total Sales"
        ];
        if (calculation is not null)
            arguments.AddRange(["--calculation", calculation]);
        var result = await InProcessCliHelper.RunAsync(arguments, _ =>
        {
            dispatched = true;
            return new ServiceResponse { Success = true };
        });
        Assert.NotEqual(0, result.ExitCode);
        Assert.False(dispatched);
    }
}

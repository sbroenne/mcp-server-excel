using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFillContractCliTests
{
    [Theory]
    [InlineData("down")]
    [InlineData("up")]
    [InlineData("left")]
    [InlineData("right")]
    public async Task Fill_MapsDirectionAndDefaultOverwritePolicy(string direction)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "fill", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1:B3", "--direction", direction
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"destinationAddress":"$A$2:$B$3"}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.fill", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(direction, args.RootElement.GetProperty("direction").GetString(), ignoreCase: true);
        Assert.False(args.RootElement.TryGetProperty("overwritePolicy", out _));
    }

    [Fact]
    public async Task AutoFill_MapsSourceDestinationAndNativeType()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "auto-fill", "--session", "session-1",
            "--sheet-name", "Sheet1", "--source-range", "A1:A2",
            "--destination-range", "A1:A10", "--fill-type", "series"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.auto-fill", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("A1:A2", args.RootElement.GetProperty("sourceRange").GetString());
        Assert.Equal("A1:A10", args.RootElement.GetProperty("destinationRange").GetString());
        Assert.Equal("series", args.RootElement.GetProperty("fillType").GetString(), ignoreCase: true);
    }

    [Fact]
    public async Task CreateSeries_MapsTypedNumericOptions()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "create-series", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1:A10",
            "--orientation", "columns", "--series-type", "linear", "--step-value", "2.5"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.create-series", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(2.5, args.RootElement.GetProperty("stepValue").GetDouble());
        Assert.Equal("columns", args.RootElement.GetProperty("orientation").GetString(), ignoreCase: true);
    }

    [Theory]
    [InlineData("get-formulas")]
    [InlineData("set-formulas")]
    public async Task Formulas_MapExplicitR1C1Notation(string action)
    {
        ServiceRequest? captured = null;
        List<string> arguments =
        [
            "range", action, "--session", "session-1", "--sheet-name", "Sheet1",
            "--range-address", "B1", "--reference-style", "r1c1"
        ];
        if (action == "set-formulas")
            arguments.AddRange(["--formulas", """[["=RC[-1]*2"]]"""]);
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("r1c1", args.RootElement.GetProperty("referenceStyle").GetString(), ignoreCase: true);
    }
}

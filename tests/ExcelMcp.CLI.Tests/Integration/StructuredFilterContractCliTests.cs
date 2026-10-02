using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "StructuredFilters")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class StructuredFilterContractCliTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ApplyFilter_MapsTheCorrectTypedOptionsParameter(bool table)
    {
        ServiceRequest? captured = null;
        List<string> arguments = ["-q", table ? "tablecolumn" : "rangeedit", "apply-filter", "--session", "session-1"];
        if (table)
            arguments.AddRange(["--table-name", "Sales", "--column-name", "Amount", "--options"]);
        else
            arguments.AddRange(["--sheet-name", "Data", "--range-address", "A1:B6", "--column-index", "2", "--filter-options"]);
        arguments.Add("""{"filterOperator":"And","criteria1":">=20","criteria2":"<=40"}""");
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal(table ? "tablecolumn.apply-filter" : "rangeedit.apply-filter", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        var options = args.RootElement.GetProperty(table ? "options" : "filterOptions");
        Assert.Equal("And", options.GetProperty("filterOperator").GetString());
        Assert.Equal(">=20", options.GetProperty("criteria1").GetString());
        Assert.Equal("<=40", options.GetProperty("criteria2").GetString());
    }

    [Fact]
    public async Task AdvancedFilter_MapsCopyModeAndExplicitPermission()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "advanced-filter", "--session", "session-1", "--sheet-name", "Data",
            "--range-address", "A1:B6", "--criteria-range", "H1:H2", "--mode", "Copy",
            "--copy-to-range", "D1", "--unique-only", "true", "--overwrite-policy", "allow"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"checkedDestinationRange":"$D$1:$E$6"}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.advanced-filter", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Copy", args.RootElement.GetProperty("mode").GetString(), ignoreCase: true);
        Assert.Equal("D1", args.RootElement.GetProperty("copyToRange").GetString());
        Assert.True(args.RootElement.GetProperty("uniqueOnly").GetBoolean());
    }

    [Fact]
    public async Task RemovedValueListAction_IsNotAnAlias()
    {
        int calls = 0;
        var result = await InProcessCliHelper.RunAsync(
        ["tablecolumn", "apply-filter-values", "--session", "session-1"], _ =>
        {
            calls++;
            return new ServiceResponse { Success = true };
        });
        Assert.NotEqual(0, result.ExitCode);
        Assert.Equal(0, calls);
    }
}

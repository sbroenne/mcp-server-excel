using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotDepth")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotDepthContractCliTests
{
    [Theory]
    [InlineData("pivottablefield", "add-field-filter", "--filter-options", """{"type":"ValueIsGreaterThan","number1":150,"dataFieldName":"Total Sales"}""", "filterOptions", "type", "ValueIsGreaterThan")]
    [InlineData("pivottablecalc", "set-layout-options", "--layout-options", """{"rowLayout":1,"repeatLabels":true,"styleName":"PivotStyleMedium9"}""", "layoutOptions", "styleName", "PivotStyleMedium9")]
    public async Task TypedOptions_KeepExactNestedNames(string category, string action, string flag, string payload,
        string nested, string property, string expected)
    {
        ServiceRequest? captured = null;
        var arguments = new List<string> { "-q", category, action, "--session", "session-1", "--pivot-table-name", "Sales", flag, payload };
        if (category == "pivottablefield")
            arguments.AddRange(["--field-name", "Region"]);
        var result = await InProcessCliHelper.RunAsync([.. arguments], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"{category}.{action}", captured.Command);
        using var document = JsonDocument.Parse(captured.Args!);
        Assert.Equal(expected, document.RootElement.GetProperty(nested).GetProperty(property).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SourceSelection_MapsExactlyOneRangeOrTable(bool table)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "pivottable", "set-source", "--session", "session-1", "--pivot-table-name", "Sales",
                "--source-sheet-name", "Data", table ? "--table-name" : "--source-range-address", table ? "Source" : "A1:C20"], request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        using var document = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Data", document.RootElement.GetProperty("sourceSheetName").GetString());
        Assert.Equal(table ? "Source" : "A1:C20", document.RootElement.GetProperty(table ? "tableName" : "sourceRangeAddress").GetString());
        Assert.False(document.RootElement.TryGetProperty(table ? "sourceRangeAddress" : "tableName", out _));
    }

    [Fact]
    public async Task ItemExpansion_ForwardsExplicitFalse()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "pivottablefield", "set-item-expansion", "--session", "session-1", "--pivot-table-name", "Sales",
                "--field-name", "Region", "--item-name", "North", "--expanded", "false"], request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        using var document = JsonDocument.Parse(captured.Args!);
        Assert.False(document.RootElement.GetProperty("expanded").GetBoolean());
    }
}

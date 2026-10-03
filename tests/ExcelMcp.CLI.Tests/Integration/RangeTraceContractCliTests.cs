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
public sealed class RangeTraceContractCliTests
{
    [Theory]
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    public async Task Trace_MapsExactScopeAndReturnsCoverage(string action)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", action, "--session", "session-1", "--sheet-name", "Sheet1",
            "--range-address", "A1,C3"
        ], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"coverage":{"scope":"same-worksheet-only","workbookComplete":false},"nodes":[],"edges":[],"unresolved":[]}"""
            };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"range.{action}", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1,C3", args.RootElement.GetProperty("rangeAddress").GetString());
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.False(output.RootElement.GetProperty("coverage").GetProperty("workbookComplete").GetBoolean());
    }
}

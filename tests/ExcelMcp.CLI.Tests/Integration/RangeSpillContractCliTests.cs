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
public sealed class RangeSpillContractCliTests
{
    [Fact]
    public async Task GetSpillInfo_MapsScopeAndPreservesCompleteRelationships()
    {
        ServiceRequest? captured = null;
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            capability = "supported",
            cellCount = 64,
            cells = Enumerable.Range(1, 64).Select(row => new
            {
                address = $"$A${row}",
                state = row == 1 ? "source" : "result",
                sourceAddress = "$A$1",
                spillAddress = "$A$1:$A$64"
            }).ToArray()
        });
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", "get-spill-info", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1:A64"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = response };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("range.get-spill-info", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1:A64", args.RootElement.GetProperty("rangeAddress").GetString());
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal("supported", output.RootElement.GetProperty("capability").GetString());
        Assert.Equal(64, output.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(64, output.RootElement.GetProperty("cells").GetArrayLength());
        Assert.Equal("$A$1", output.RootElement.GetProperty("cells")[63].GetProperty("sourceAddress").GetString());
    }
}

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
public sealed class RangeSpecialCellsContractCliTests
{
    [Theory]
    [InlineData("formulas")]
    [InlineData("constants")]
    [InlineData("blanks")]
    [InlineData("errors")]
    [InlineData("visible")]
    public async Task SpecialCells_MapsSelectorAndPreservesCompleteResult(string cellKind)
    {
        ServiceRequest? captured = null;
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            sheetName = "Sheet1",
            rangeAddress = "$A$1:$A$64",
            cellKind,
            cellCount = 32,
            areas = Enumerable.Range(0, 32).Select(index => $"$A${index * 2 + 1}").ToArray()
        });
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", "get-special-cells", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1:A64", "--cell-kind", cellKind
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = response };
        });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("range.get-special-cells", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1:A64", args.RootElement.GetProperty("rangeAddress").GetString());
        Assert.Equal(cellKind, args.RootElement.GetProperty("cellKind").GetString(), ignoreCase: true);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(32, output.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(32, output.RootElement.GetProperty("areas").GetArrayLength());
    }

    [Theory]
    [InlineData("unknown-kind")]
    [InlineData("99")]
    public async Task SpecialCells_InvalidSelectorDoesNotDispatch(string cellKind)
    {
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", "get-special-cells", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1", "--cell-kind", cellKind
        ]);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("cellKind", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SpecialCells_MissingSelectorDoesNotDispatch()
    {
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", "get-special-cells", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1"
        ]);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("cellKind", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }
}

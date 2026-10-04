using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class RangeFormulaErrorCliParityTests : IAsyncLifetime
{
    private static readonly string[][] ReferenceErrorFormula = [["=INDIRECT(\"A0\")"]];

    private readonly CliWorkbookSessionFixture _workbook = new();

    public Task InitializeAsync() => _workbook.InitializeAsync();
    public Task DisposeAsync() => _workbook.DisposeAsync();

    [Fact]
    public async Task RangeReads_ReturnCanonicalFormulaErrorThroughCli()
    {
        string sessionId = _workbook.SessionId;

        string formulas = JsonSerializer.Serialize(ReferenceErrorFormula);
        var set = await CliProcessHelper.RunAsync(
            [
                "range", "set-formulas",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1",
                    "--formulas", formulas
            ]);
        Assert.Equal(0, set.ExitCode);
        using var setDocument = JsonDocument.Parse(set.Stdout);
        Assert.True(setDocument.RootElement.GetProperty("success").GetBoolean());

        var values = await CliProcessHelper.RunAsync(
            [
                "range", "get-values",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1"
            ]);
        var formulasResult = await CliProcessHelper.RunAsync(
            [
                "range", "get-formulas",
                    "--session", sessionId!,
                    "--sheet-name", "Sheet1",
                    "--range-address", "A1"
            ]);

        Assert.Equal(0, values.ExitCode);
        Assert.Equal(0, formulasResult.ExitCode);
        AssertCanonicalReferenceError(values.Stdout);
        AssertCanonicalReferenceError(formulasResult.Stdout);
    }

    private static void AssertCanonicalReferenceError(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal("#REF!", Assert.Single(Assert.Single(root.GetProperty("values").EnumerateArray()).EnumerateArray()).GetString());

        var error = Assert.Single(root.GetProperty("cellErrors").EnumerateArray());
        Assert.Equal("A1", error.GetProperty("cellAddress").GetString());
        Assert.Equal("#REF!", error.GetProperty("errorName").GetString());
        Assert.Equal("=INDIRECT(\"A0\")", error.GetProperty("formula").GetString());
        Assert.Equal(-2146826265, error.GetProperty("errorCode").GetInt32());
        Assert.Equal(-2146826265, error.GetProperty("currentValue").GetInt32());
    }

}

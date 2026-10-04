using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "DataModel")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class DataModelFormulaExcelParityTests : IAsyncLifetime
{
    private readonly CliWorkbookSessionFixture _workbook = new();

    public Task InitializeAsync() => _workbook.InitializeAsync();

    public Task DisposeAsync() => _workbook.DisposeAsync();

    [Fact]
    public async Task DaxMeasureWrites_NativeCommaSyntax_PreserveAndEvaluate()
    {
        await RunAsync(
        [
            "powerquery", "create", "--session", _workbook.SessionId,
            "--query-name", "SalesTable",
            "--m-code", "#table(type table [Amount = number], {{1000}, {2500}})",
            "--load-destination", "data-model"
        ]);
        const string formula = "DIVIDE(SUM(SalesTable[Amount]), 1000)";
        foreach (var update in new[] { false, true })
        {
            var name = $"Comma_{Guid.NewGuid():N}";
            await RunAsync(
            [
                "datamodel", "create-measure", "--session", _workbook.SessionId,
                "--table-name", "SalesTable", "--measure-name", name,
                "--dax-formula", update ? "SUM(SalesTable[Amount])" : formula,
                "--format-type", update ? "Percentage" : "Decimal"
            ]);
            if (update)
            {
                await RunAsync(
                [
                    "datamodel", "update-measure", "--session", _workbook.SessionId,
                    "--measure-name", name, "--dax-formula", formula,
                    "--format-type", "Decimal"
                ]);
            }
            var read = await RunAsync(
            [
                "datamodel", "read", "--session", _workbook.SessionId, "--measure-name", name
            ]);
            Assert.Equal(formula, read.GetProperty("daxFormula").GetString());
            Assert.Equal(name, read.GetProperty("measureName").GetString());
            Assert.Equal("SalesTable", read.GetProperty("tableName").GetString());
            Assert.Equal(formula.Length, read.GetProperty("characterCount").GetInt32());
            Assert.Equal("Decimal", read.GetProperty("formatInfo").GetProperty("type").GetString());
            var evaluated = await RunAsync(
            [
                "datamodel", "evaluate", "--session", _workbook.SessionId,
                "--dax-query", $"EVALUATE ROW(\"Result\", [{name}])"
            ]);
            Assert.Equal(3.5m, Assert.Single(Assert.Single(
                evaluated.GetProperty("rows").EnumerateArray()).EnumerateArray()).GetDecimal());
        }
    }

    private static async Task<JsonElement> RunAsync(IReadOnlyList<string> arguments)
    {
        var result = await CliProcessHelper.RunAsync(arguments, timeoutMs: 120000);
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        using var document = JsonDocument.Parse(result.Stdout);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
        if (document.RootElement.TryGetProperty("errorMessage", out var error))
        {
            Assert.True(error.ValueKind == JsonValueKind.Null || string.IsNullOrEmpty(error.GetString()), result.Stdout);
        }
        return document.RootElement.Clone();
    }
}

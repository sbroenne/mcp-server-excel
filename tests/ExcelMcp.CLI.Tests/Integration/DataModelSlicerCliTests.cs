using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class DataModelSlicerCliTests : IAsyncLifetime
{
    private readonly CliWorkbookSessionFixture _workbook = new();
    private static readonly string[] Captions = ["2026 Q1", "2026 Q2", "2026 Q3"];

    public Task InitializeAsync() => _workbook.InitializeAsync();
    public Task DisposeAsync() => _workbook.DisposeAsync();

    [Fact]
    public async Task DataModelSlicer_ExecutableAndPipe_FilterRecoveryAndClearChangePivotValues()
    {
        await RunAsync("sheet", "create", "--sheet-name", "SlicerSmoke");
        await RunAsync("powerquery", "create", "--query-name", "SmokeSales",
            "--m-code", "#table(type table [Quarter = text, Amount = number], {{\"2026 Q1\", 10}, {\"2026 Q2\", 20}, {\"2026 Q2\", 30}, {\"2026 Q3\", 40}})",
            "--load-destination", "data-model");
        await RunAsync("datamodel", "create-measure", "--table-name", "SmokeSales",
            "--measure-name", "SmokeTotal", "--dax-formula", "SUM(SmokeSales[Amount])");
        await RunAsync("pivottable", "create-from-datamodel", "--table-name", "SmokeSales",
            "--destination-sheet", "SlicerSmoke", "--destination-cell", "A1",
            "--pivot-table-name", "SmokePivot");
        await RunAsync("pivottablefield", "add-row-field", "--pivot-table-name", "SmokePivot",
            "--field-name", "[SmokeSales].[Quarter]");
        await RunAsync("pivottablefield", "add-value-field", "--pivot-table-name", "SmokePivot",
            "--field-name", "[Measures].[SmokeTotal]");
        await RunAsync("pivottable", "refresh", "--pivot-table-name", "SmokePivot");
        var created = await RunAsync("slicer", "create-slicer", "--pivot-table-name", "SmokePivot",
            "--field-name", "[SmokeSales].[Quarter]", "--slicer-name", "SmokeSlicer",
            "--destination-sheet", "SlicerSmoke", "--position", "D1");
        AssertItems(created, Captions);
        Assert.Equal("SmokePivot", Assert.Single(created.GetProperty("connectedPivotTables")
            .EnumerateArray()).GetString());
        await AssertStateAsync(100, Captions);

        var selected = await RunAsync("slicer", "set-slicer-selection", "--slicer-name", "SmokeSlicer",
            "--selected-items", "[\"2026 Q2\"]");
        AssertItems(selected, [Captions[1]]);
        await AssertStateAsync(50, Captions[1]);

        var arguments = WithSession("slicer", "set-slicer-selection", "--slicer-name", "SmokeSlicer",
            "--selected-items", "[\"2026 Q1\",\"missing\"]");
        var failed = await CliProcessHelper.RunAsync(arguments, timeoutMs: 120000);
        Assert.NotEqual(0, failed.ExitCode);
        using (var error = JsonDocument.Parse(failed.Stdout))
        {
            Assert.False(error.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("was not found", error.RootElement.GetProperty("errorMessage").GetString());
        }
        await AssertStateAsync(50, Captions[1]);

        // Omitting clear-first must replace, not add to, the existing Q2 filter.
        AssertItems(await RunAsync("slicer", "set-slicer-selection", "--slicer-name", "SmokeSlicer",
            "--selected-items", "[\"2026 Q1\"]"), [Captions[0]]);
        await AssertStateAsync(10, Captions[0]);
        AssertItems(await RunAsync("slicer", "set-slicer-selection", "--slicer-name", "SmokeSlicer",
            "--selected-items", "[\"2026 Q2\"]", "--clear-first", "false"), [Captions[0], Captions[1]]);
        await AssertStateAsync(60, Captions[0], Captions[1]);
        AssertItems(await RunAsync("slicer", "set-slicer-selection", "--slicer-name", "SmokeSlicer",
            "--selected-items", "[]", "--clear-first", "false"), Captions);
        await AssertStateAsync(100, Captions);
    }

    private async Task AssertStateAsync(double total, params string[] selected)
    {
        var listed = await RunAsync("slicer", "list-slicers", "--pivot-table-name", "SmokePivot");
        AssertItems(Assert.Single(listed.GetProperty("slicers").EnumerateArray()), selected);
        var values = await RunAsync("range", "get-values", "--sheet-name", "SlicerSmoke",
            "--range-address", "A1:B6");
        var rows = values.GetProperty("values").EnumerateArray()
            .Where(row => row[1].ValueKind == JsonValueKind.Number).ToArray();
        Assert.Equal(selected.Length + 1, rows.Length);
        Assert.Equal(selected.Order(), rows.Take(selected.Length).Select(row => row[0].GetString()).Order());
        foreach (var row in rows.Take(selected.Length))
        {
            Assert.Equal(row[0].GetString() switch
            {
                "2026 Q1" => 10d,
                "2026 Q2" => 50d,
                "2026 Q3" => 40d,
                _ => throw new InvalidOperationException("Unexpected PivotTable caption.")
            }, row[1].GetDouble());
        }
        Assert.Equal(total, rows[^1][1].GetDouble());
    }

    private static void AssertItems(JsonElement slicer, string[] selected)
    {
        Assert.Equal(Captions, slicer.GetProperty("availableItems")
            .EnumerateArray().Select(item => item.GetString()).Order());
        Assert.Equal(selected.Order(), slicer.GetProperty("selectedItems")
            .EnumerateArray().Select(item => item.GetString()).Order());
    }

    private string[] WithSession(params string[] arguments) =>
        [.. arguments, "--session", _workbook.SessionId];

    private async Task<JsonElement> RunAsync(params string[] arguments)
    {
        var result = await CliProcessHelper.RunAsync(WithSession(arguments), timeoutMs: 120000);
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        using var document = JsonDocument.Parse(result.Stdout);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean(), result.Stdout);
        Assert.True(!document.RootElement.TryGetProperty("errorMessage", out var error)
            || error.ValueKind == JsonValueKind.Null || error.GetString() == string.Empty, result.Stdout);
        return document.RootElement.Clone();
    }
}

using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class RangeFindExcelParityTests : IAsyncLifetime
{
    private readonly CliWorkbookSessionFixture _workbook = new();

    public Task InitializeAsync() => _workbook.InitializeAsync();

    public Task DisposeAsync() => _workbook.DisposeAsync();

    [Fact]
    public async Task Find_ReturnsExactCoverageThroughExecutableAndExcel()
    {
        var values = Enumerable.Range(0, 26)
            .Select(index => new[] { index < 25 ? "Apple" : "Banana" }).ToArray();
        var set = await CliProcessHelper.RunAsync(
        [
            "range", "set-values", "--session", _workbook.SessionId,
            "--sheet-name", "Sheet1", "--range-address", "A1:A26",
            "--values", JsonSerializer.Serialize(values)
        ]);
        Assert.True(set.ExitCode == 0, set.Stdout + set.Stderr);
        using var setJson = JsonDocument.Parse(set.Stdout);
        Assert.True(setJson.RootElement.GetProperty("success").GetBoolean());

        foreach (int? limit in new int?[] { null, 5 })
        {
            List<string> arguments =
            [
                "rangeedit", "find", "--session", _workbook.SessionId,
                "--sheet-name", "Sheet1", "--range-address", "A1:A26",
                "--search-value", "Apple", "--find-options", """{"matchEntireCell":true}"""
            ];
            if (limit.HasValue)
            {
                arguments.Add("--max-matches");
                arguments.Add(limit.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            var find = await CliProcessHelper.RunAsync(arguments);
            Assert.True(find.ExitCode == 0, find.Stdout + find.Stderr);
            using var document = JsonDocument.Parse(find.Stdout);
            var root = document.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean());
            Assert.Equal(25, root.GetProperty("totalCount").GetInt64());
            Assert.Equal(limit ?? 10, root.GetProperty("returnedCount").GetInt32());
            Assert.Equal(limit ?? 10, root.GetProperty("matchingCells").GetArrayLength());
            Assert.True(root.GetProperty("truncated").GetBoolean());
            var matches = root.GetProperty("matchingCells").EnumerateArray().ToArray();
            Assert.Equal(matches.Length, matches.Select(cell => cell.GetProperty("address").GetString()).Distinct().Count());
            foreach (var cell in matches)
            {
                int row = cell.GetProperty("row").GetInt32();
                Assert.InRange(row, 1, 25);
                Assert.Equal(1, cell.GetProperty("column").GetInt32());
                Assert.Equal($"$A${row}", cell.GetProperty("address").GetString());
                Assert.Equal("Apple", cell.GetProperty("value").GetString());
            }
        }
    }
}

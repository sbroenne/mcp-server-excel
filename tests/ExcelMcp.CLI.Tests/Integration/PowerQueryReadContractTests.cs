using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class PowerQueryReadContractTests : IAsyncLifetime
{
    private readonly CliWorkbookSessionFixture _workbook = new();

    public Task InitializeAsync() => _workbook.InitializeAsync();
    public Task DisposeAsync() => _workbook.DisposeAsync();

    [Fact]
    public async Task List_IsCompactAndViewReturnsFullM_ViaCliProtocol()
    {
        const string queryName = "CliCompactRead";
        var mCode = BuildLongMCode();
        var (createResult, createJson) = await CliProcessHelper.RunJsonAsync(
            [
                "powerquery", "create",
                "--session", _workbook.SessionId,
                "--query-name", queryName,
                "--m-code", mCode,
                "--load-destination", "connection-only"
            ],
            timeoutMs: 120_000,
            diagnosticLabel: "pq-read-contract-create");

        Assert.Equal(0, createResult.ExitCode);
        Assert.True(createJson.RootElement.GetProperty("success").GetBoolean());

        var (listResult, listJson) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "list", "--session", _workbook.SessionId],
            timeoutMs: 60_000,
            diagnosticLabel: "pq-read-contract-list");

        Assert.Equal(0, listResult.ExitCode);
        Assert.True(listJson.RootElement.GetProperty("success").GetBoolean());
        Assert.True(listResult.Stdout.Length < 1_000);
        var serializedQuery = Assert.Single(
            listJson.RootElement.GetProperty("queries").EnumerateArray(),
            item => item.GetProperty("name").GetString() == queryName);
        Assert.False(serializedQuery.TryGetProperty("formula", out _));
        Assert.InRange(serializedQuery.GetProperty("formulaPreview").GetString()!.Length, 1, 80);
        Assert.Equal(mCode.Length, serializedQuery.GetProperty("characterCount").GetInt32());
        Assert.Equal("connection-only", serializedQuery.GetProperty("loadMode").GetString());

        var (viewResult, viewJson) = await CliProcessHelper.RunJsonAsync(
            [
                "powerquery", "view",
                "--session", _workbook.SessionId,
                "--query-name", queryName
            ],
            timeoutMs: 60_000,
            diagnosticLabel: "pq-read-contract-view");

        Assert.Equal(0, viewResult.ExitCode);
        Assert.True(viewJson.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(mCode, viewJson.RootElement.GetProperty("mCode").GetString());
        Assert.Equal("connection-only", viewJson.RootElement.GetProperty("loadMode").GetString());
    }

    private static string BuildLongMCode()
    {
        var padding = string.Join(
            Environment.NewLine,
            Enumerable.Repeat("// CLI list must not serialize this padding", 250));
        return $"let{Environment.NewLine}{padding}{Environment.NewLine}    Source = #table({{\"Value\"}}, {{{{1}}}}){Environment.NewLine}in{Environment.NewLine}    Source";
    }
}

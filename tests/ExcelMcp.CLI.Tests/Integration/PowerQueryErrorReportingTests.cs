using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class PowerQueryErrorReportingTests : IAsyncLifetime
{
    private readonly CliWorkbookSessionFixture _workbook = new();

    public Task InitializeAsync() => _workbook.InitializeAsync();
    public Task DisposeAsync() => _workbook.DisposeAsync();

    [Fact]
    public async Task Refresh_SyntheticFirewallError_ReturnsStructuredPrivacyCategory()
    {
        var queryName = "SyntheticFirewallQuery";
        var validMCode = """
            let
                Source = #table({"X"}, {{1}})
            in
                Source
            """;
        var firewallMCode = """
            let
                Root = error Error.Record(
                    "Formula.Firewall",
                    "Query 'ConfigData' (step 'Root') references other queries or steps, so it may not directly access a data source.",
                    null)
            in
                Root
            """;

        var (createResult, createJson) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "create", "--session", _workbook.SessionId, "--query-name", queryName, "--m-code", validMCode],
            timeoutMs: 120000,
            diagnosticLabel: "pq-error-reporting-create");

        Assert.Equal(0, createResult.ExitCode);
        Assert.True(createJson.RootElement.GetProperty("success").GetBoolean());
        await AssertLoadedValueAsync(queryName, 1);

        var (updateResult, updateJson) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "update", "--session", _workbook.SessionId, "--query-name", queryName, "--m-code", firewallMCode, "--refresh", "false"],
            timeoutMs: 120000,
            diagnosticLabel: "pq-error-reporting-update");

        Assert.Equal(0, updateResult.ExitCode);
        Assert.True(updateJson.RootElement.GetProperty("success").GetBoolean());
        await AssertLoadedValueAsync(queryName, 1);

        var (refreshResult, refreshJson) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "refresh", "--session", _workbook.SessionId, "--query-name", queryName, "--timeout-seconds", "60"],
            timeoutMs: 120000,
            diagnosticLabel: "pq-error-reporting-refresh");

        Assert.NotEqual(0, refreshResult.ExitCode);
        Assert.False(refreshJson.RootElement.GetProperty("success").GetBoolean());
        Assert.True(refreshJson.RootElement.GetProperty("isError").GetBoolean());
        Assert.Equal("Privacy", refreshJson.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal("PowerQueryCommandException", refreshJson.RootElement.GetProperty("exceptionType").GetString());
        Assert.Equal(
            refreshJson.RootElement.GetProperty("error").GetString(),
            refreshJson.RootElement.GetProperty("errorMessage").GetString());
        Assert.Equal("0x800A03EC", refreshJson.RootElement.GetProperty("hresult").GetString());
        AssertOptionalNonEmptyStringProperty(refreshJson.RootElement, "innerError");
        Assert.Contains("Formula.Firewall", refreshJson.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
        await AssertLoadedValueAsync(queryName, 1);
        var (viewResult, viewJson) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "view", "--session", _workbook.SessionId, "--query-name", queryName]);
        using (viewJson)
        {
            Assert.Equal(0, viewResult.ExitCode);
            Assert.True(viewJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(firewallMCode, viewJson.RootElement.GetProperty("mCode").GetString());
        }

        var (recoveryResult, recoveryJson) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "update", "--session", _workbook.SessionId, "--query-name", queryName,
                    "--m-code", "#table({\"X\"}, {{2}})"],
            timeoutMs: 120000,
            diagnosticLabel: "pq-error-reporting-recovery");
        using (recoveryJson)
        {
            Assert.Equal(0, recoveryResult.ExitCode);
            Assert.True(recoveryJson.RootElement.GetProperty("success").GetBoolean());
        }
        await AssertLoadedValueAsync(queryName, 2);
    }

    private async Task AssertLoadedValueAsync(string sheetName, int expected)
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["range", "get-values", "--session", _workbook.SessionId, "--sheet-name", sheetName,
                "--range-address", "A1:A2"]);
        using (json)
        {
            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
            var rows = json.RootElement.GetProperty("values");
            Assert.Equal(2, rows.GetArrayLength());
            Assert.Equal("X", Assert.Single(rows[0].EnumerateArray()).GetString());
            Assert.Equal(expected, Assert.Single(rows[1].EnumerateArray()).GetInt32());
        }
    }

    private static void AssertOptionalNonEmptyStringProperty(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out var property))
        {
            return;
        }

        Assert.False(string.IsNullOrWhiteSpace(property.GetString()), $"{propertyName} should be non-empty when present.");
    }
}

using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class PowerQueryErrorReportingTests : IDisposable
{
    private readonly string _testFile = Path.Combine(Path.GetTempPath(), $"PqErrorReporting_{Guid.NewGuid():N}.xlsx");
    private string? _sessionId;

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

        var (sessionResult, sessionJson) = await CliProcessHelper.RunJsonAsync(
            ["session", "create", _testFile],
            timeoutMs: 60000,
            diagnosticLabel: "pq-error-reporting-session-create");

        Assert.Equal(0, sessionResult.ExitCode);
        _sessionId = sessionJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(_sessionId));

        try
        {
            var (createResult, createJson) = await CliProcessHelper.RunJsonAsync(
                ["powerquery", "create", "--session", _sessionId!, "--query-name", queryName, "--m-code", validMCode],
                timeoutMs: 120000,
                diagnosticLabel: "pq-error-reporting-create");

            Assert.Equal(0, createResult.ExitCode);
            Assert.True(createJson.RootElement.GetProperty("success").GetBoolean());

            var (updateResult, updateJson) = await CliProcessHelper.RunJsonAsync(
                ["powerquery", "update", "--session", _sessionId!, "--query-name", queryName, "--m-code", firewallMCode, "--refresh", "false"],
                timeoutMs: 120000,
                diagnosticLabel: "pq-error-reporting-update");

            Assert.Equal(0, updateResult.ExitCode);
            Assert.True(updateJson.RootElement.GetProperty("success").GetBoolean());

            var (refreshResult, refreshJson) = await CliProcessHelper.RunJsonAsync(
                ["powerquery", "refresh", "--session", _sessionId!, "--query-name", queryName, "--timeout", "60"],
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
            Assert.False(refreshJson.RootElement.TryGetProperty("hresult", out _));
            AssertOptionalNonEmptyStringProperty(refreshJson.RootElement, "innerError");
            Assert.Contains("Formula.Firewall", refreshJson.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(_sessionId))
            {
#pragma warning disable CA1031
                try
                {
                    await CliProcessHelper.RunAsync(
                        ["session", "close", "--session", _sessionId!, "--save", "false"],
                        timeoutMs: 60000,
                        diagnosticLabel: "pq-error-reporting-close");
                }
                catch
                {
                }
#pragma warning restore CA1031
            }
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFile))
        {
            try
            {
                File.Delete(_testFile);
            }
            catch
            {
            }
        }

        GC.SuppressFinalize(this);
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

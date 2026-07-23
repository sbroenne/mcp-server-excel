using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PowerQueryErrorReportingProtocolTests : McpIntegrationTestBase
{
    private readonly string _testExcelFile;
    private string? _sessionId;

    public PowerQueryErrorReportingProtocolTests(ITestOutputHelper output)
        : base(output, "PowerQueryErrorReportingProtocolClient")
    {
        _testExcelFile = Path.Join(CreateTempDirectory("PowerQueryErrorReporting"), "PowerQueryErrorReporting.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_testExcelFile);
    }

    [Fact]
    public async Task Refresh_SyntheticFirewallError_ReturnsStructuredDiagnosticsViaMcpProtocol()
    {
        const string queryName = "SyntheticFirewallQuery";
        const string validMCode = """
            let
                Source = #table({"X"}, {{1}})
            in
                Source
            """;
        const string firewallMCode = """
            let
                Root = error Error.Record(
                    "Formula.Firewall",
                    "Query 'ConfigData' (step 'Root') references other queries or steps, so it may not directly access a data source.",
                    null)
            in
                Root
            """;

        var createQueryJson = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["query_name"] = queryName,
            ["m_code"] = validMCode
        });
        AssertSuccess(createQueryJson, "powerquery.create");

        var updateQueryJson = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "update",
            ["session_id"] = _sessionId,
            ["query_name"] = queryName,
            ["m_code"] = firewallMCode,
            ["refresh"] = false
        });
        AssertSuccess(updateQueryJson, "powerquery.update");

        var refreshJson = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "refresh",
            ["session_id"] = _sessionId,
            ["query_name"] = queryName,
            ["timeout"] = 60
        });

        using var doc = ParseJsonResult(refreshJson, "powerquery.refresh synthetic-firewall");
        AssertFailureEnvelope(
            doc.RootElement,
            "powerquery.refresh synthetic-firewall",
            expectedExceptionType: "PowerQueryCommandException",
            expectedErrorCategory: "Privacy",
            allowOptionalNonEmptyInnerError: true);

        Assert.Contains("Formula.Firewall", doc.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
    }
}

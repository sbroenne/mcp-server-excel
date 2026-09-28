using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class PowerQueryErrorReportingProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task Refresh_SyntheticFirewallError_ReturnsStructuredDiagnosticsViaMcpProtocol()
    {
        const string sessionId = "recording-session";
        const string queryName = "SyntheticFirewallQuery";
        var call = await _fixture.CallToolAsync(
            "powerquery",
            new Dictionary<string, object?>
            {
                ["action"] = "refresh",
                ["session_id"] = sessionId,
                ["query_name"] = queryName,
                ["timeout_seconds"] = 60
            },
            new ServiceResponse
            {
                Success = false,
                Command = "powerquery.refresh",
                SessionId = sessionId,
                ErrorMessage =
                    "Formula.Firewall: Query 'ConfigData' references other queries.",
                ExceptionType = "PowerQueryCommandException",
                ErrorCategory = "Privacy",
                HResult = "0x800A03EC",
                InnerError = "Formula.Firewall"
            },
            "powerquery.refresh",
            """{"queryName":"SyntheticFirewallQuery","timeout":60}""");

        using (var args = RecordingToolTest.ParseArgs(
            call.Request,
            "powerquery.refresh",
            sessionId))
        {
            Assert.Equal(
                queryName,
                args.RootElement.GetProperty("queryName").GetString());
            Assert.Equal(
                60,
                args.RootElement.GetProperty("timeout").GetInt32());
        }

        using var document = JsonDocument.Parse(call.JsonResult);
        var root = document.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(
            "PowerQueryCommandException",
            root.GetProperty("exceptionType").GetString());
        Assert.Equal("Privacy", root.GetProperty("errorCategory").GetString());
        Assert.Equal("0x800A03EC", root.GetProperty("hresult").GetString());
        if (root.TryGetProperty("innerError", out var innerError))
        {
            Assert.False(string.IsNullOrWhiteSpace(innerError.GetString()));
        }
        Assert.Contains(
            "Formula.Firewall",
            root.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);
    }
}

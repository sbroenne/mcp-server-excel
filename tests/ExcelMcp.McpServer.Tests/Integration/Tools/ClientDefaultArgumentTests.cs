using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Some MCP clients fill in every optional input with its advertised default (issue #1089).
/// </summary>
[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class ClientDefaultArgumentTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task ManualTool_UnusedInputsWithAdvertisedDefaults_AreAccepted()
    {
        var call = await fixture.CallToolAsync(
            "file_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["file_path"] = null,
                ["timeout_seconds"] = 120
            },
            new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" },
            "session.list",
            expectedSessionId: null,
            expectedArgsJson: null);

        Assert.False(call.Result.IsError == true);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task GeneratedTool_UnusedInputsWithNullDefaults_AreAccepted()
    {
        var call = await fixture.CallToolAsync(
            "calculation_mode",
            new Dictionary<string, object?>
            {
                ["action"] = "calculate",
                ["workbook_session_id"] = "session-1",
                ["scope"] = "application",
                ["mode"] = null,
                ["iteration_enabled"] = null,
                ["precision_as_displayed"] = null,
                ["allow_precision_loss"] = null
            },
            new ServiceResponse { Success = true, Result = """{"success":true}""" },
            "calculationmode.calculate",
            JsonSerializer.Serialize(new { scope = "application" }, ServiceProtocol.JsonOptions));

        Assert.False(call.Result.IsError == true);
    }
}

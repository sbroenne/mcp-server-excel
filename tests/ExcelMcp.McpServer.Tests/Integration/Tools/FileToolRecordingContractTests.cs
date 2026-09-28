using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "false")]
public sealed class FileToolRecordingContractTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Theory]
    [InlineData("create", "session.create", false)]
    [InlineData("open", "session.open", true)]
    public async Task CreateAndOpen_DispatchNullSessionAndExplicitDefaults(
        string action,
        string expectedCommand,
        bool createExistingFile)
    {
        var path = Path.Join(
            Path.GetTempPath(),
            $"file-recording-{Guid.NewGuid():N}.xlsx");
        if (createExistingFile)
        {
            await File.WriteAllTextAsync(path, string.Empty);
        }

        try
        {
            var response = new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    sessionId = "recorded-session",
                    filePath = path
                }, ServiceProtocol.JsonOptions)
            };
            var expectedArgs = action == "create"
                ? JsonSerializer.Serialize(new
                {
                    filePath = path,
                    macroEnabled = false,
                    show = false,
                    timeoutSeconds = 120
                }, ServiceProtocol.JsonOptions)
                : JsonSerializer.Serialize(new
                {
                    filePath = path,
                    show = false,
                    timeoutSeconds = 120
                }, ServiceProtocol.JsonOptions);

            var call = await _fixture.CallToolAsync(
                "file",
                new Dictionary<string, object?>
                {
                    ["action"] = action,
                    ["path"] = path,
                    ["save"] = false,
                    ["show"] = false,
                    ["timeout_seconds"] = 120
                },
                response,
                expectedCommand,
                expectedSessionId: null,
                expectedArgsJson: expectedArgs);

            Assert.Null(call.Request.SessionId);
            using var result = JsonDocument.Parse(call.JsonResult);
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(
                "recorded-session",
                result.RootElement.GetProperty("session_id").GetString());
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task List_DispatchesNullSessionAndNullArguments()
    {
        var response = new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"sessions":[]}"""
        };

        var call = await _fixture.CallToolAsync(
            "file",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["save"] = false,
                ["show"] = false,
                ["timeout_seconds"] = 120
            },
            response,
            "session.list",
            expectedSessionId: null,
            expectedArgsJson: null);

        Assert.Null(call.Request.SessionId);
        Assert.Null(call.Request.Args);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(0, result.RootElement.GetProperty("sessions").GetArrayLength());
    }

    [Fact]
    public async Task Close_DispatchesExactSessionAndSaveDefault()
    {
        var call = await _fixture.CallToolAsync(
            "file",
            new Dictionary<string, object?>
            {
                ["action"] = "close",
                ["session_id"] = "session-close",
                ["save"] = false,
                ["show"] = false,
                ["timeout_seconds"] = 120
            },
            new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"session_id":"session-close","saved":false}"""
            },
            "session.close",
            "session-close",
            """{"save":false}""");

        Assert.Equal("session-close", call.Request.SessionId);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("saved").GetBoolean());
    }
}

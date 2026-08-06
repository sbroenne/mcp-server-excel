using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "SessionPreflight")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class SessionPreflightCliTests(
    ITestOutputHelper output,
    TempDirectoryFixture fixture) : IClassFixture<TempDirectoryFixture>
{
    [Fact]
    public async Task SessionPreflight_OpenSession_ReturnsSameCapabilityContractAsMcp()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            "Preflight",
            "contract",
            fixture.TempDir,
            ".xlsx");
        string? sessionId = null;

        try
        {
            var (createResult, createJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "open", workbookPath],
                timeoutMs: 30000,
                diagnosticLabel: "session-preflight-create");
            using var createJson = createJsonDocument;
            Assert.Equal(0, createResult.ExitCode);
            sessionId = createJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));

            var (preflightResult, preflightJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "preflight", "--session-id", sessionId!],
                timeoutMs: 30000,
                diagnosticLabel: "session-preflight");
            using var preflightJson = preflightJsonDocument;

            output.WriteLine($"[session-preflight] Stdout: {preflightResult.Stdout}");
            output.WriteLine($"[session-preflight] Stderr: {preflightResult.Stderr}");

            Assert.Equal(0, preflightResult.ExitCode);
            Assert.True(preflightJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(sessionId, preflightJson.RootElement.GetProperty("sessionId").GetString());
            Assert.Equal(workbookPath, preflightJson.RootElement.GetProperty("filePath").GetString());
            Assert.True(preflightJson.RootElement.GetProperty("excel").GetProperty("build").GetInt32() > 0);
            Assert.True(new[] { "supported", "unsupported" }.Contains(
                preflightJson.RootElement.GetProperty("capabilities").GetProperty("formula2").GetProperty("status").GetString()));
            Assert.Equal("notDetermined", preflightJson.RootElement.GetProperty("capabilities").GetProperty("pythonInExcel").GetProperty("status").GetString());
            Assert.Equal(JsonValueKind.Array, preflightJson.RootElement.GetProperty("constraints").ValueKind);
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                await CliProcessHelper.RunAsync(
                    ["session", "close", "--session", sessionId],
                    timeoutMs: 30000,
                    diagnosticLabel: "session-preflight-close");
            }
        }
    }
}

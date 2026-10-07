// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end regressions for file tool behavior through the MCP protocol.
/// These tests use the real transport and server pipeline instead of calling tool methods directly.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "true")]
public sealed class ExcelFileToolProtocolRegressionTests : McpIntegrationTestBase
{
    public ExcelFileToolProtocolRegressionTests(ITestOutputHelper output)
        : base(output, "ExcelFileToolProtocolRegressionClient")
    {
    }

    private static string? GetConfiguredIrmTestFilePath()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        return !string.IsNullOrWhiteSpace(irmTestFile) && File.Exists(irmTestFile)
            ? Path.GetFullPath(irmTestFile)
            : null;
    }

    [ConfiguredIrmFact]
    [Trait("RunType", "OnDemand")]
    public async Task FileOpen_RealIrmWorkbook_ReturnsWithinTimeoutBudget_WhenConfigured()
    {
        // Real IRM/AIP workbooks require local auth state and are intentionally opt-in only.
        var irmTestFile = GetConfiguredIrmTestFilePath()
            ?? throw new InvalidOperationException("Configured IRM test fixture was unavailable after test discovery.");

        var testResult = await CallToolAsync("file_read", new Dictionary<string, object?>
        {
            ["action"] = "test",
            ["file_path"] = irmTestFile
        });

        using (var testJson = JsonDocument.Parse(testResult))
        {
            Assert.True(testJson.RootElement.GetProperty("success").GetBoolean());
            Assert.True(testJson.RootElement.GetProperty("isIrmProtected").GetBoolean());
        }

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        var openResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["file_path"] = irmTestFile,
            ["timeout_seconds"] = 15
        }).WaitAsync(TimeSpan.FromSeconds(20));
        stopwatch.Stop();

        Output.WriteLine($"IRM open result after {stopwatch.Elapsed.TotalSeconds:F1}s: {openResult}");

        using var openJson = JsonDocument.Parse(openResult);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(20),
            "MCP file.open must return within the requested timeout budget for protected workbooks.");
        Assert.True(openJson.RootElement.TryGetProperty("success", out var successProp));

        string? sessionId = null;
        if (successProp.GetBoolean())
        {
            sessionId = openJson.RootElement.GetProperty("workbook_session_id").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));
        }
        else
        {
            var errorMessage = openJson.RootElement.GetProperty("errorMessage").GetString();
            Assert.False(string.IsNullOrWhiteSpace(errorMessage));
        }

        var listResult = await CallToolAsync("file_read", new Dictionary<string, object?>
        {
            ["action"] = "list"
        });

        using (var listJson = JsonDocument.Parse(listResult))
        {
            Assert.True(listJson.RootElement.GetProperty("success").GetBoolean());
        }

        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            TrackSession(sessionId);
            await CloseSessionAsync(sessionId, save: false);
        }
    }

}

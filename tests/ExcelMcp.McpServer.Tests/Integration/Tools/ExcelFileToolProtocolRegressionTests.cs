// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
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
    private readonly string _tempDir;

    public ExcelFileToolProtocolRegressionTests(ITestOutputHelper output)
        : base(output, "ExcelFileToolProtocolRegressionClient")
    {
        _tempDir = CreateTempDirectory("ExcelFileToolProtocolRegressionTests");
    }

    private static string? GetConfiguredIrmTestFilePath()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        return !string.IsNullOrWhiteSpace(irmTestFile) && File.Exists(irmTestFile)
            ? Path.GetFullPath(irmTestFile)
            : null;
    }

    [Fact]
    public async Task FileOpen_FileLockedByAnotherProcess_ReturnsActionableError_AndNextOpenSucceeds()
    {
        var lockedFile = Path.Join(_tempDir, $"LockedOpen_{Guid.NewGuid():N}.xlsx");
        ExcelSession.CreateNew<bool>(lockedFile, false, (ctx, ct) => true);

        using (var fileLock = new FileStream(lockedFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var lockedResult = await CallToolAsync("file", new Dictionary<string, object?>
            {
                ["action"] = "open",
                ["path"] = lockedFile
            });

            Output.WriteLine($"Locked file open result: {lockedResult}");

            using var lockedJson = JsonDocument.Parse(lockedResult);
            Assert.False(lockedJson.RootElement.GetProperty("success").GetBoolean());
            Assert.True(lockedJson.RootElement.GetProperty("isError").GetBoolean());

            var errorMessage = lockedJson.RootElement.GetProperty("errorMessage").GetString();
            Assert.NotNull(errorMessage);
            Assert.Contains("already open", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("close the file", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exclusive access", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.False(string.IsNullOrWhiteSpace(lockedJson.RootElement.GetProperty("exceptionType").GetString()));
        }

        var listAfterFailure = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "list"
        });

        using (var listAfterFailureJson = JsonDocument.Parse(listAfterFailure))
        {
            Assert.True(listAfterFailureJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(0, listAfterFailureJson.RootElement.GetProperty("sessions").GetArrayLength());
        }

        var openAfterRelease = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = lockedFile
        });
        AssertSuccess(openAfterRelease, "Open workbook after lock release");

        var sessionId = GetJsonProperty(openAfterRelease, "session_id");
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        TrackSession(sessionId);

        await CloseSessionAsync(sessionId, save: false);
        await Task.Delay(TimeSpan.FromSeconds(2));
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task FileOpen_RealIrmWorkbook_ReturnsWithinTimeoutBudget_WhenConfigured()
    {
        // Real IRM/AIP workbooks require local auth state and are intentionally opt-in only.
        var irmTestFile = GetConfiguredIrmTestFilePath();
        if (irmTestFile == null)
        {
            return;
        }

        var testResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "test",
            ["path"] = irmTestFile
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
            ["path"] = irmTestFile,
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
            sessionId = openJson.RootElement.GetProperty("session_id").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));
        }
        else
        {
            var errorMessage = openJson.RootElement.GetProperty("errorMessage").GetString();
            Assert.False(string.IsNullOrWhiteSpace(errorMessage));
        }

        var listResult = await CallToolAsync("file", new Dictionary<string, object?>
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

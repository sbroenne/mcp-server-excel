// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end smoke tests for the MCP Server using the official MCP SDK client.
///
/// PURPOSE: Validates the complete MCP protocol stack works correctly with real Excel operations.
/// PATTERN: Uses Program.RunAsync with host-owned in-memory transport and the real service.
/// Each scenario owns its prerequisites and workbook (requires Excel COM automation).
///
/// These tests exercise:
/// - Full DI pipeline (exact same as production)
/// - MCP protocol serialization/deserialization
/// - Tool discovery and invocation via MCP protocol
/// - Real Excel operations through COM interop
/// - Session management across multiple tool calls
/// - Application Insights telemetry (same configuration as production)
///
/// The server is a BLACK BOX - tests only interact via MCP protocol.
/// Only the transport differs: pipes instead of stdio.
///
/// Run before commits to catch breaking changes:
/// dotnet test --filter "FullyQualifiedName~McpServerSmokeTests"
/// </summary>
[Collection("ProgramTransport")]  // Real Excel tests must run sequentially.
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SmokeTest")]
[Trait("RequiresExcel", "true")]
public class McpServerSmokeTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;

    // MCP transport pipes
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;
    private bool _disposed;

    public McpServerSmokeTests(ITestOutputHelper output)
    {
        _output = output;

        // Create temp directory for test files
        _tempDir = Path.Join(Path.GetTempPath(), $"McpSmokeTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Join(_tempDir, "SmokeTest.xlsx");

        _output.WriteLine($"Test directory: {_tempDir}");
    }

    /// <summary>
    /// Setup: Configure test transport and run the real MCP server.
    /// The server is a BLACK BOX - we only configure transport, everything else is production code.
    /// </summary>
    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            _cts.Token,
            "SmokeTestClient");

        _output.WriteLine($"✓ Connected to server: {_client.ServerInfo?.Name} v{_client.ServerInfo?.Version}");
    }

    public async Task DisposeAsync()
    {
        await DisposeAsyncCore();
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        if (_disposed)
            return;
        _disposed = true;

        var failures = new List<Exception>();
        try
        {
            await ProgramTransportTestHost.StopAsync(
                _client,
                _clientToServerPipe,
                _serverToClientPipe,
                _serverTask,
                _output);
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            _cts.Dispose();
            try
            {
                if (Directory.Exists(_tempDir))
                    Directory.Delete(_tempDir, recursive: true);
            }
            catch (Exception cleanupFailure)
            {
                failures.Add(cleanupFailure);
            }
        }
        if (failures.Count != 0)
            throw new AggregateException("MCP shutdown or file cleanup failed.", failures);
    }

    /// <summary>
    /// Comprehensive smoke test that exercises the MCP tools via the SDK client.
    /// This validates the complete E2E flow: MCP protocol → DI → Tool → Core → Excel COM.
    /// </summary>
    [Fact]
    [Trait("Acceptance", "Required")]
    public async Task Smoke_WorkbookLifecycle_PersistsValues()
    {
        _output.WriteLine("=== MCP SERVER E2E SMOKE TEST (SDK CLIENT) ===");
        _output.WriteLine("Testing selected operations via MCP protocol with real Excel...\n");

        // =====================================================================
        // STEP 1: CREATE AND OPEN SESSION
        // =====================================================================
        _output.WriteLine("✓ Step 1: Creating workbook and opening session via MCP protocol...");

        var createResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["file_path"] = _testExcelFile
        });
        AssertSuccess(createResult, "File creation and session open");
        Assert.True(File.Exists(_testExcelFile), "Excel file should exist");
        var sessionId = GetJsonProperty(createResult, "workbook_session_id");
        Assert.NotNull(sessionId);
        _output.WriteLine($"  ✓ file: Create passed (session: {sessionId})");

        var listSessionsResult = await CallToolAsync("file_read", new Dictionary<string, object?>
        {
            ["action"] = "list"
        });
        AssertSuccess(listSessionsResult, "List workbook sessions");
        using (var listed = JsonDocument.Parse(listSessionsResult))
        {
            var session = Assert.Single(listed.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(sessionId, session.GetProperty("sessionId").GetString());
            Assert.False(session.TryGetProperty("workbook_session_id", out _));
            Assert.True(session.GetProperty("canClose").GetBoolean());
        }
        await CallSuccessfulToolAsync("worksheet", new()
        {
            ["action"] = "create",
            ["workbook_session_id"] = sessionId,
            ["sheet_name"] = "Data"
        });
        await CallSuccessfulToolAsync("range", new()
        {
            ["action"] = "set-values",
            ["workbook_session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1",
            ["values"] = new List<List<string>> { new() { "persisted-smoke-value" } }
        });
        var selected = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["workbook_session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:B2"
        });
        AssertSuccess(selected, "Cell read");
        using (var read = JsonDocument.Parse(selected))
        {
            var root = read.RootElement;
            Assert.Equal(2, root.GetProperty("rowCount").GetInt32());
            Assert.Equal(2, root.GetProperty("columnCount").GetInt32());
            Assert.Equal("persisted-smoke-value", root.GetProperty("values")[0][0].GetString());
            Assert.Equal(JsonValueKind.Null, root.GetProperty("values")[1][1].ValueKind);
            Assert.False(root.TryGetProperty("filePath", out _));
        }
        await SaveAndVerifyAsync(sessionId, "persisted-smoke-value");
    }


    private async Task SaveAndVerifyAsync(string sessionId, string expectedValue)
    {
        // =====================================================================
        // STEP 14: CLOSE SESSION (save changes)
        // =====================================================================
        _output.WriteLine("\n✓ Step 14: Closing session (saving changes)...");

        var closeResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["workbook_session_id"] = sessionId,
            ["save"] = true
        });
        AssertSuccess(closeResult, "Close session");
        _output.WriteLine("  ✓ Session saved and closed");

        // =====================================================================
        // STEP 15: VERIFY PERSISTENCE
        // =====================================================================
        _output.WriteLine("\n✓ Step 15: Verifying persistence...");

        var verifyOpenResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["file_path"] = _testExcelFile
        });
        AssertSuccess(verifyOpenResult, "Re-open for verification");
        var verifySessionId = GetJsonProperty(verifyOpenResult, "workbook_session_id");
        Assert.NotNull(verifySessionId);

        await VerifyAndCloseAsync(verifySessionId, async () =>
        {
            var finalSheetsResult = await CallToolAsync("worksheet_read", new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["workbook_session_id"] = verifySessionId
            });
            AssertSuccess(finalSheetsResult, "Final worksheet list");

            using (var sheetsJson = JsonDocument.Parse(finalSheetsResult))
            {
                Assert.Contains(sheetsJson.RootElement.GetProperty("worksheets").EnumerateArray(),
                    sheet => sheet.GetProperty("name").GetString() == "Data");
            }
            var values = await CallSuccessfulToolAsync("range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = verifySessionId,
                ["sheet_name"] = "Data",
                ["range_address"] = "A1"
            });
            Assert.Equal(expectedValue, GetFirstCellValue(values));
            _output.WriteLine("  Saved worksheet value survived reopening.");
        });

    }

    private Task<string> CloseSmokeWorkbookAsync(string session) =>
        CallSuccessfulToolAsync("file", new()
        {
            ["action"] = "close",
            ["workbook_session_id"] = session,
            ["save"] = false
        });

    private async Task<string> CallSuccessfulToolAsync(string tool, Dictionary<string, object?> arguments)
    {
        var response = await CallToolAsync(tool, arguments);
        AssertSuccess(response, $"{tool}.{arguments["action"]}");
        return response;
    }

    private async Task VerifyAndCloseAsync(string sessionId, Func<Task> verify)
    {
        var failures = new List<Exception>();
        try
        {
            await verify();
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            try
            {
                await CloseSmokeWorkbookAsync(sessionId);
            }
            catch (Exception cleanupFailure)
            {
                failures.Add(cleanupFailure);
            }
        }
        if (failures.Count != 0)
            throw new AggregateException("MCP verification or session close failed.", failures);
    }

    /// <summary>
    /// Calls a tool via the MCP protocol and returns the text response.
    /// </summary>
    private async Task<string> CallToolAsync(
        string toolName,
        Dictionary<string, object?> arguments,
        bool expectedError = false)
    {
        var result = await _client!.CallToolAsync(toolName, arguments, cancellationToken: _cts.Token);
        return McpResponseAssertions.ReadText(result, expectedError);
    }

    /// <summary>
    /// Asserts the JSON response indicates success.
    /// </summary>
    private static void AssertSuccess(string jsonResult, string operationName) =>
        McpResponseAssertions.AssertSuccess(jsonResult, operationName);

    private static string? GetFirstCellValue(string jsonResult)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement
            .GetProperty("values")[0][0]
            .GetString();
    }

    /// <summary>
    /// Gets a string property from a JSON response.
    /// </summary>
    private static string? GetJsonProperty(string jsonResult, string propertyName)
    {
        using var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) ? prop.GetString() : null;
    }
}

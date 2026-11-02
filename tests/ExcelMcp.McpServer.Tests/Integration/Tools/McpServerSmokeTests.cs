// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
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
    private static readonly string[] StartupCommands = ["session.create", "sheet.create", "range.set-values"];
    private static readonly string[][] InitialValues = [["persisted-smoke-value"]];
    private static readonly string[][] DefaultSessionValues = [["default-session"]];
    private static readonly string[][] OverrideSessionValues = [["override-session"]];
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;
    private readonly HashSet<string> _trackedSessionIds = new(StringComparer.Ordinal);

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
        foreach (var sessionId in _trackedSessionIds.ToArray())
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

        var createResult = await CallSuccessfulToolAsync("batch", new()
        {
            ["action"] = "run",
            ["commands"] = new object[]
            {
                new { command = "session.create", args = new { filePath = _testExcelFile } },
                new { command = "sheet.create", args = new { sheetName = "Data" } },
                new
                {
                    command = "range.set-values",
                    args = new { sheetName = "Data", rangeAddress = "A1", values = InitialValues }
                }
            }
        });
        Assert.True(File.Exists(_testExcelFile), "Excel file should exist");
        using var created = JsonDocument.Parse(createResult);
        AssertBatchCounts(created.RootElement, total: 3, processed: 3);
        var startupResults = created.RootElement.GetProperty("results").EnumerateArray().ToArray();
        Assert.Equal(StartupCommands,
            startupResults.Select(result => result.GetProperty("command").GetString()));
        Assert.All(startupResults, result => Assert.True(result.GetProperty("success").GetBoolean()));
        var sessionId = startupResults[0].GetProperty("workbook_session_id").GetString();
        Assert.NotNull(sessionId);
        _output.WriteLine($"  Batch startup captured its session for worksheet creation and writing: {sessionId}");

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
        await VerifyBatchFailuresAsync(sessionId);
        await VerifyBatchSessionOverrideAsync(sessionId);
        await SaveAndVerifyAsync(sessionId,
        [
            ["persisted-smoke-value", "before-failure", "after-failure", "default-session"],
            [null, "before-stop", null, null]
        ]);
    }

    private async Task VerifyBatchFailuresAsync(string sessionId)
    {
        foreach (var stopOnError in new[] { false, true })
        {
            var row = stopOnError ? 2 : 1;
            var before = stopOnError ? "before-stop" : "before-failure";
            var arguments = new Dictionary<string, object?>
            {
                ["action"] = "run",
                ["workbook_session_id"] = sessionId,
                ["commands"] = new object[]
                {
                    new
                    {
                        command = "range.set-values",
                        args = new { sheetName = "Data", rangeAddress = $"B{row}", values = new[] { new[] { before } } }
                    },
                    new { command = "range.get-values", args = new { sheetName = "MissingBatchSheet", rangeAddress = "A1" } },
                    new
                    {
                        command = "range.set-values",
                        args = new { sheetName = "Data", rangeAddress = $"C{row}", values = new[] { new[] { "after-failure" } } }
                    }
                }
            };
            if (stopOnError)
                arguments["stop_on_error"] = true;
            var json = await CallToolAsync("batch", arguments, expectedError: true);
            using var batch = JsonDocument.Parse(json);
            var root = batch.RootElement;
            Assert.False(root.GetProperty("success").GetBoolean());
            AssertBatchCounts(root, total: 3, processed: stopOnError ? 2 : 3);
            var results = root.GetProperty("results");
            Assert.True(results[0].GetProperty("success").GetBoolean());
            Assert.False(results[1].GetProperty("success").GetBoolean());
            Assert.Contains("MissingBatchSheet", results[1].GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
            if (!stopOnError)
                Assert.True(results[2].GetProperty("success").GetBoolean());

            var read = await CallSuccessfulToolAsync("range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = sessionId,
                ["sheet_name"] = "Data",
                ["range_address"] = $"B{row}:C{row}"
            });
            using var cells = JsonDocument.Parse(read);
            AssertValues(cells.RootElement, [[before, stopOnError ? null : "after-failure"]]);
        }
    }

    private async Task VerifyBatchSessionOverrideAsync(string sessionId)
    {
        var other = await CallSuccessfulToolAsync("file", new()
        {
            ["action"] = "create",
            ["file_path"] = Path.Join(_tempDir, "OtherWorkbook.xlsx")
        });
        var otherSessionId = GetJsonProperty(other, "workbook_session_id");
        Assert.NotNull(otherSessionId);
        await VerifyAndCloseAsync(otherSessionId, async () =>
        {
            await CallSuccessfulToolAsync("worksheet", new()
            {
                ["action"] = "create",
                ["workbook_session_id"] = otherSessionId,
                ["sheet_name"] = "Data"
            });
            var json = await CallSuccessfulToolAsync("batch", new()
            {
                ["action"] = "run",
                ["workbook_session_id"] = sessionId,
                ["commands"] = new object[]
                {
                    new
                    {
                        command = "range.set-values",
                        args = new { sheetName = "Data", rangeAddress = "D1", values = DefaultSessionValues }
                    },
                    new
                    {
                        command = "range.set-values",
                        sessionId = otherSessionId,
                        args = new { sheetName = "Data", rangeAddress = "D1", values = OverrideSessionValues }
                    },
                    new { command = "range.get-values", args = new { sheetName = "Data", rangeAddress = "D1" } }
                }
            });
            using var batch = JsonDocument.Parse(json);
            AssertBatchCounts(batch.RootElement, total: 3, processed: 3);
            Assert.All(batch.RootElement.GetProperty("results").EnumerateArray(),
                result => Assert.True(result.GetProperty("success").GetBoolean()));
            AssertValues(batch.RootElement.GetProperty("results")[2], [["default-session"]]);
            var read = await CallSuccessfulToolAsync("range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = otherSessionId,
                ["sheet_name"] = "Data",
                ["range_address"] = "D1"
            });
            using var cells = JsonDocument.Parse(read);
            AssertValues(cells.RootElement, [["override-session"]]);
        });
    }

    private async Task SaveAndVerifyAsync(string sessionId, string?[][] expectedValues)
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
        using var closed = JsonDocument.Parse(closeResult);
        var timings = closed.RootElement.GetProperty("closeTimingsMs");
        Assert.Equal(11, timings.EnumerateObject().Count());
        Assert.True(timings.GetProperty("session-close").GetDouble() > 0);
        Assert.True(timings.GetProperty("save").GetDouble() > 0);
        Assert.True(timings.GetProperty("workbook-close").GetDouble() > 0);
        Assert.True(timings.GetProperty("worker-completion").GetDouble() >= 0);
        Assert.All(timings.EnumerateObject(), timing => Assert.True(timing.Value.GetDouble() >= 0));
        Assert.True(System.Text.Encoding.UTF8.GetByteCount(timings.GetRawText()) < 1024);
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
                ["range_address"] = "A1:D2"
            });
            using var cells = JsonDocument.Parse(values);
            AssertValues(cells.RootElement, expectedValues);
            _output.WriteLine("  Exact batch results and untouched cells survived reopening.");
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
        using var json = JsonDocument.Parse(Assert.Single(result.Content.OfType<TextContentBlock>()).Text);
        if (toolName == "batch" && json.RootElement.TryGetProperty("results", out var items))
        {
            foreach (var item in items.EnumerateArray().Where(item =>
                item.GetProperty("command").GetString() is "session.create" or "session.open" or "session.attach" &&
                item.GetProperty("success").GetBoolean()))
                _trackedSessionIds.Add(item.GetProperty("workbook_session_id").GetString()!);
        }
        if (toolName == "file" && json.RootElement.GetProperty("success").GetBoolean())
        {
            if (arguments["action"] is "create" or "open")
                _trackedSessionIds.Add(json.RootElement.GetProperty("workbook_session_id").GetString()!);
            if (arguments["action"] is "close" && arguments["workbook_session_id"] is string closedSessionId)
                _trackedSessionIds.Remove(closedSessionId);
        }
        return McpResponseAssertions.ReadText(result, expectedError);
    }

    /// <summary>
    /// Asserts the JSON response indicates success.
    /// </summary>
    private static void AssertSuccess(string jsonResult, string operationName) =>
        McpResponseAssertions.AssertSuccess(jsonResult, operationName);

    private static void AssertBatchCounts(JsonElement root, int total, int processed)
    {
        Assert.Equal(total, root.GetProperty("totalCount").GetInt32());
        Assert.Equal(processed, root.GetProperty("executedCount").GetInt32());
        Assert.Equal(processed < total, root.GetProperty("stoppedEarly").GetBoolean());
        var results = root.GetProperty("results").EnumerateArray().ToArray();
        Assert.Equal(processed, results.Length);
        Assert.Equal(Enumerable.Range(0, processed), results.Select(result => result.GetProperty("index").GetInt32()));
    }

    private static void AssertValues(JsonElement root, string?[][] expected)
    {
        Assert.Equal(expected.Length, root.GetProperty("rowCount").GetInt32());
        Assert.Equal(expected[0].Length, root.GetProperty("columnCount").GetInt32());
        Assert.True(JsonElement.DeepEquals(JsonSerializer.SerializeToElement(expected), root.GetProperty("values")),
            $"Expected {JsonSerializer.Serialize(expected)}, got {root.GetProperty("values").GetRawText()}.");
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

// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end smoke tests for the MCP Server using the official MCP SDK client.
///
/// PURPOSE: Validates the complete MCP protocol stack works correctly with real Excel operations.
/// PATTERN: Uses Program.ConfigureTestTransport() to inject in-memory pipes, then runs the real server.
/// RUNTIME: ~30-60 seconds (requires Excel COM automation).
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
[Collection("ProgramTransport")]  // Uses Program.ConfigureTestTransport() - must run sequentially
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
    private readonly string _testCsvFile;

    // MCP transport pipes
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpServerSmokeTests(ITestOutputHelper output)
    {
        _output = output;

        // Create temp directory for test files
        _tempDir = Path.Join(Path.GetTempPath(), $"McpSmokeTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Join(_tempDir, "SmokeTest.xlsx");
        _testCsvFile = Path.Join(_tempDir, "SampleData.csv");

        _output.WriteLine($"Test directory: {_tempDir}");
    }

    /// <summary>
    /// Setup: Configure test transport and run the real MCP server.
    /// The server is a BLACK BOX - we only configure transport, everything else is production code.
    /// </summary>
    public async Task InitializeAsync()
    {
        // Configure the server to use our test pipes instead of stdio
        // This is the ONLY difference from production - transport layer only
        Program.ConfigureTestTransport(_clientToServerPipe, _serverToClientPipe);

        // Run the REAL server (Program.Main) - exact same code path as production
        // The server will use our configured pipes for transport
        _serverTask = Program.Main([]);

        // Allow server to initialize before client connection
        // SDK 0.5.0+ has stricter initialization timing
        await Task.Delay(100);

        // Create client connected to the server via pipes
        _client = await McpClient.CreateAsync(
            new StreamClientTransport(
                serverInput: _clientToServerPipe.Writer.AsStream(),
                serverOutput: _serverToClientPipe.Reader.AsStream()),
            clientOptions: new McpClientOptions
            {
                ClientInfo = new() { Name = "SmokeTestClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)  // Increase timeout for test stability
            },
            cancellationToken: _cts.Token);

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
        // Flush telemetry before shutdown to ensure test telemetry is sent
        ExcelMcpTelemetry.Flush();

        // Dispose client first - this signals we're done sending requests
        if (_client != null)
        {
            await _client.DisposeAsync();
        }

        // Complete the pipes to signal EOF - this triggers GRACEFUL server shutdown
        // The MCP SDK will see EOF and stop the host naturally, allowing
        // Application Insights and other services to flush during shutdown
        _clientToServerPipe.Writer.Complete();
        _serverToClientPipe.Writer.Complete();

        // Wait for server to shut down gracefully (with timeout)
        if (_serverTask != null)
        {
            // Give the server time to flush telemetry and clean up
            var shutdownTimeout = Task.Delay(TimeSpan.FromSeconds(10));
            var completed = await Task.WhenAny(_serverTask, shutdownTimeout);

            if (completed == shutdownTimeout)
            {
                // Server didn't shut down in time - cancel as fallback
                await _cts.CancelAsync();
                try
                {
                    await _serverTask;
                }
                catch (OperationCanceledException)
                {
                    // Expected when we had to force cancel
                }
            }
        }

        // Reset test transport for next test
        Program.ResetTestTransport();

        _cts.Dispose();

        // Clean up temp files
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }

    /// <summary>
    /// Comprehensive smoke test that exercises all 12 MCP tools via the SDK client.
    /// This validates the complete E2E flow: MCP protocol → DI → Tool → Core → Excel COM.
    /// </summary>
    [Fact]
    public async Task SmokeTest_AllTools_E2EWorkflow()
    {
        _output.WriteLine("=== MCP SERVER E2E SMOKE TEST (SDK CLIENT) ===");
        _output.WriteLine("Testing all 25 tools via MCP protocol with real Excel...\n");

        // =====================================================================
        // STEP 1: CREATE AND OPEN SESSION
        // =====================================================================
        _output.WriteLine("✓ Step 1: Creating workbook and opening session via MCP protocol...");

        var createResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = _testExcelFile
        });
        AssertSuccess(createResult, "File creation and session open");
        Assert.True(File.Exists(_testExcelFile), "Excel file should exist");
        var sessionId = GetJsonProperty(createResult, "session_id");
        Assert.NotNull(sessionId);
        _output.WriteLine($"  ✓ file: Create passed (session: {sessionId})");

        // =====================================================================
        // STEP 3: WORKSHEET OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 3: Worksheet operations...");

        var listSheetsResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listSheetsResult, "List worksheets");

        var createSheetResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data"
        });
        AssertSuccess(createSheetResult, "Create worksheet");
        _output.WriteLine("  ✓ worksheet: List and Create passed");

        // =====================================================================
        // STEP 4: RANGE OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 4: Range operations...");

        var values = new List<List<object?>>
        {
            new() { "Name", "Value", "Date" },
            new() { "Item1", 100, "2024-01-01" },
            new() { "Item2", 200, "2024-01-02" }
        };

        var setValuesResult = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C3",
            ["values"] = values
        });
        AssertSuccess(setValuesResult, "Set values");

        var getValuesResult = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C3"
        });
        AssertSuccess(getValuesResult, "Get values");
        _output.WriteLine("  ✓ range: SetValues and GetValues passed");

        // =====================================================================
        // STEP 5: TABLE OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 5: Table operations...");

        var createTableResult = await CallToolAsync("table", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId,
            ["table_name"] = "DataTable",
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C3",
            ["has_headers"] = true
        });
        AssertSuccess(createTableResult, "Create table");

        var listTablesResult = await CallToolAsync("table", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId
        });
        AssertSuccess(listTablesResult, "List tables");
        _output.WriteLine("  ✓ table: Create and List passed");

        // =====================================================================
        // STEP 6: NAMED RANGE OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 6: Named range operations...");

        var createParamResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId,
            ["name"] = "ReportDate",
            ["reference"] = "=Data!$C$2"
        });
        AssertSuccess(createParamResult, "Create named range");

        var readParamResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "read",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId,
            ["name"] = "ReportDate"
        });
        AssertSuccess(readParamResult, "Read named range");
        _output.WriteLine("  ✓ namedrange: Create and Read passed");

        // =====================================================================
        // STEP 7: POWER QUERY OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 7: Power Query operations...");

        // Create test CSV
        var csvContent = "Product,Quantity\nWidget,10\nGadget,20";
        await File.WriteAllTextAsync(_testCsvFile, csvContent);

        var mCode = $@"let
    Source = Csv.Document(File.Contents(""{_testCsvFile.Replace("\\", "\\\\")}""),[Delimiter="","", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders";

        var createQueryResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["query_name"] = "CsvData",
            ["m_code"] = mCode,
            ["load_destination"] = "connection-only"
        });
        AssertSuccess(createQueryResult, "Create Power Query");

        var listQueriesResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listQueriesResult, "List Power Queries");

        // Rename the query (US1: Power Query rename)
        var renameQueryResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = sessionId,
            ["old_name"] = "CsvData",
            ["new_name"] = "ProductData"
        });
        AssertSuccess(renameQueryResult, "Rename Power Query");
        Assert.Contains("ProductData", renameQueryResult);

        // Verify rename by listing again
        var listAfterRenameResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listAfterRenameResult, "List Power Queries after rename");
        Assert.Contains("ProductData", listAfterRenameResult);
        Assert.DoesNotContain("CsvData", listAfterRenameResult);

        _output.WriteLine("  ✓ powerquery: Create, List, and Rename passed");

        // =====================================================================
        // STEP 8: CONNECTION OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 8: Connection operations...");

        var listConnectionsResult = await CallToolAsync("connection", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId
        });
        AssertSuccess(listConnectionsResult, "List connections");
        _output.WriteLine("  ✓ connection: List passed");

        // =====================================================================
        // STEP 9: PIVOTTABLE OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 9: PivotTable operations...");

        var createPivotResult = await CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "create-from-table",
            ["session_id"] = sessionId,
            ["table_name"] = "DataTable",
            ["destination_sheet"] = "Data",
            ["destination_cell"] = "E1",
            ["pivot_table_name"] = "SalesPivot"
        });
        AssertSuccess(createPivotResult, "Create PivotTable");

        var listPivotsResult = await CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listPivotsResult, "List PivotTables");
        _output.WriteLine("  ✓ pivottable: Create and List passed");

        // =====================================================================
        // STEP 10: CHART OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 10: Chart operations...");

        var createChartResult = await CallToolAsync("chart", new Dictionary<string, object?>
        {
            ["action"] = "create-from-range",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["source_range_address"] = "A1:C3",
            ["chart_type"] = "ColumnClustered",
            ["left"] = 50,
            ["top"] = 50,
            ["width"] = 400,
            ["height"] = 300,
            ["chart_name"] = "DataChart"
        });
        AssertSuccess(createChartResult, "Create Chart");

        var listChartsResult = await CallToolAsync("chart", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        // Chart List returns array directly
        Assert.NotNull(listChartsResult);
        _output.WriteLine("  ✓ chart: Create and List passed");

        // =====================================================================
        // STEP 11: DATA MODEL OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 11: Data Model operations...");

        var listDataModelResult = await CallToolAsync("datamodel", new Dictionary<string, object?>
        {
            ["action"] = "list-tables",
            ["session_id"] = sessionId
        });
        AssertSuccess(listDataModelResult, "List Data Model tables");

        // Test rename-table returns expected failure due to Excel limitation (not a crash)
        // First, we need a PQ-backed table in the Data Model
        // The ProductData query was created above - load it to Data Model
        var loadToDmResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "load-to",
            ["session_id"] = sessionId,
            ["query_name"] = "ProductData",
            ["load_destination"] = "load-to-data-model"
        });
        AssertSuccess(loadToDmResult, "Load Power Query to Data Model");

        // Verify table exists
        var listAfterLoadResult = await CallToolAsync("datamodel", new Dictionary<string, object?>
        {
            ["action"] = "list-tables",
            ["session_id"] = sessionId
        });
        AssertSuccess(listAfterLoadResult, "List Data Model tables after load");
        Assert.Contains("ProductData", listAfterLoadResult);

        // Attempt rename-table - this will return success=false due to Excel limitation
        var renameTableResult = await CallToolAsync("datamodel", new Dictionary<string, object?>
        {
            ["action"] = "rename-table",
            ["session_id"] = sessionId,
            ["old_name"] = "ProductData",
            ["new_name"] = "RenamedProductData"
        });
        // Expect JSON with success=false (not a crash)
        var renameJson = JsonDocument.Parse(renameTableResult);
        Assert.True(renameJson.RootElement.TryGetProperty("success", out var renameSuccess));
        Assert.False(renameSuccess.GetBoolean(), "Rename-table should fail due to Excel limitation");
        Assert.True(renameJson.RootElement.TryGetProperty("errorMessage", out var renameError));
        var renameErrorText = renameError.GetString() ?? "";
        // Error could be "immutable", "cannot be renamed", or "not found" (Power Query issue)
        Assert.True(
            renameErrorText.Contains("immutable", StringComparison.OrdinalIgnoreCase) ||
            renameErrorText.Contains("cannot be renamed", StringComparison.OrdinalIgnoreCase) ||
            renameErrorText.Contains("not found", StringComparison.OrdinalIgnoreCase),
            $"Expected error about rename limitation but got: {renameErrorText}");
        _output.WriteLine("  ✓ datamodel: RenameTable correctly returns error (Excel limitation)");

        _output.WriteLine("  ✓ datamodel: ListTables passed");

        // =====================================================================
        // STEP 12: CONDITIONAL FORMAT OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 12: Conditional Format operations...");

        var addRuleResult = await CallToolAsync("conditionalformat", new Dictionary<string, object?>
        {
            ["action"] = "add-rule",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "B2:B3",
            ["rule_type"] = "cellvalue",  // Note: no hyphen - Core expects "cellvalue" not "cell-value"
            ["operator_type"] = "greater",
            ["formula1"] = "100",
            ["interior_color"] = "#00FF00"
        });
        AssertSuccess(addRuleResult, "Add conditional format rule");
        _output.WriteLine("  ✓ conditionalformat: AddRule passed");

        // =====================================================================
        // STEP 13: VBA OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 13: VBA operations...");

        var listVbaResult = await CallToolAsync("vba", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["path"] = _testExcelFile,
            ["session_id"] = sessionId
        });
        AssertSuccess(listVbaResult, "List VBA modules");
        _output.WriteLine("  ✓ vba: List passed");

        // =====================================================================
        // STEP 14: CLOSE SESSION (save changes)
        // =====================================================================
        _output.WriteLine("\n✓ Step 14: Closing session (saving changes)...");

        var closeResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = sessionId,
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
            ["path"] = _testExcelFile
        });
        AssertSuccess(verifyOpenResult, "Re-open for verification");
        var verifySessionId = GetJsonProperty(verifyOpenResult, "session_id");

        try
        {
            var finalSheetsResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["session_id"] = verifySessionId
            });
            AssertSuccess(finalSheetsResult, "Final worksheet list");

            // Verify Data sheet exists
            Assert.Contains("Data", finalSheetsResult);
            _output.WriteLine("  ✓ All changes persisted correctly");
        }
        finally
        {
            await CallToolAsync("file", new Dictionary<string, object?>
            {
                ["action"] = "close",
                ["session_id"] = verifySessionId,
                ["save"] = false
            });
        }

        // =====================================================================
        // FINAL SUMMARY
        // =====================================================================
        _output.WriteLine("\n=== E2E SMOKE TEST COMPLETE ===");
        _output.WriteLine("✅ All 12 MCP tools tested via SDK client");
        _output.WriteLine("✅ Full MCP protocol stack validated");
        _output.WriteLine("✅ DI pipeline exercised (same as Program.cs)");
        _output.WriteLine("✅ Real Excel operations verified");
        _output.WriteLine("✅ Data persistence confirmed");
        _output.WriteLine("\n🚀 MCP Server E2E functionality working correctly!");
    }

    /// <summary>
    /// Tests that invalid actions return helpful error messages via MCP protocol.
    /// </summary>
    [Fact]
    public async Task InvalidSession_ReturnsHelpfulErrorMessage()
    {
        _output.WriteLine("Testing error handling via MCP protocol...");

        var result = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = "nonexistent-session-id"
        });

        _output.WriteLine($"Result: {result[..Math.Min(300, result.Length)]}...");

        // Should have success=false
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("success", out var success));
        Assert.False(success.GetBoolean());

        // Should have helpful error message
        Assert.True(json.RootElement.TryGetProperty("errorMessage", out var errorMessage));
        var errorText = errorMessage.GetString();
        Assert.NotNull(errorText);
        Assert.Contains("not found", errorText, StringComparison.OrdinalIgnoreCase);

        _output.WriteLine("✓ Error message is clear and helpful via MCP protocol");
    }

    /// <summary>
    /// Tests that worksheet copy-to-file (atomic operation) works WITHOUT session_id.
    /// This verifies the fix for the issue where copy-to-file incorrectly required session_id.
    ///
    /// Atomic operations like copy-to-file and move-to-file should NOT require a session_id
    /// because they manage their own Excel instances internally.
    /// </summary>
    [Fact]
    public async Task WorksheetCopyToFile_WithoutSessionId_Works()
    {
        _output.WriteLine("\n✓ Testing atomic worksheet.copy-to-file (no session required)...");

        // Create source and target Excel files for copying
        var sourceFile = Path.Join(_tempDir, "CopySource.xlsx");
        var targetFile = Path.Join(_tempDir, "CopyTarget.xlsx");

        // Step 1: Create source file with a sheet (use CreateNew for new files)
        _output.WriteLine("  1. Creating source file...");
        ExcelSession.CreateNew<bool>(sourceFile, false, (ctx, ct) => true);

        // Step 2: Create target file (empty, will receive the copied sheet)
        _output.WriteLine("  2. Creating target file...");
        ExcelSession.CreateNew<bool>(targetFile, false, (ctx, ct) => true);

        // Step 3: Call worksheet copy-to-file WITHOUT session_id (ATOMIC OPERATION)
        // This is the CRITICAL TEST: the tool should accept this call without a session_id parameter
        _output.WriteLine("  3. Calling worksheet.copy-to-file without session_id...");
        var copyResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "copy-to-file",
            ["source_file"] = sourceFile,
            ["source_sheet"] = "Sheet1",
            ["target_file"] = targetFile,
            ["target_sheet_name"] = "CopiedSheet"
            // NOTE: session_id is NOT provided - this is the test point!
            // Before the fix, this would fail with "sessionId is required"
        });

        AssertSuccess(copyResult, "worksheet.copy-to-file");
        _output.WriteLine("  ✓ copy-to-file succeeded WITHOUT session_id!");
        _output.WriteLine("✓ Atomic operation (copy-to-file) correctly works without session requirement!");
    }

    /// <summary>
    /// Calls a tool via the MCP protocol and returns the text response.
    /// </summary>
    private async Task<string> CallToolAsync(string toolName, Dictionary<string, object?> arguments)
    {
        var result = await _client!.CallToolAsync(toolName, arguments, cancellationToken: _cts.Token);

        Assert.NotNull(result);
        Assert.NotNull(result.Content);
        Assert.NotEmpty(result.Content);

        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        Assert.NotNull(textBlock);

        return textBlock.Text;
    }

    /// <summary>
    /// Asserts the JSON response indicates success.
    /// </summary>
    private static void AssertSuccess(string jsonResult, string operationName)
    {
        Assert.NotNull(jsonResult);

        try
        {
            var json = JsonDocument.Parse(jsonResult);

            // Check for error property
            if (json.RootElement.TryGetProperty("error", out var error))
            {
                var errorMsg = error.GetString();
                Assert.Fail($"{operationName} failed with error: {errorMsg}");
            }

            // Check for Success property (PascalCase)
            if (json.RootElement.TryGetProperty("Success", out var successPascal))
            {
                if (!successPascal.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("ErrorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned Success=false: {errorMsg}");
                }
            }
            // Check for success property (camelCase)
            else if (json.RootElement.TryGetProperty("success", out var successCamel))
            {
                if (!successCamel.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("errorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned success=false: {errorMsg}");
                }
            }
        }
        catch (JsonException ex)
        {
            Assert.Fail($"{operationName} returned invalid JSON: {ex.Message}\nResponse: {jsonResult}");
        }
    }

    /// <summary>
    /// Gets a string property from a JSON response.
    /// </summary>
    private static string? GetJsonProperty(string jsonResult, string propertyName)
    {
        var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) ? prop.GetString() : null;
    }
}






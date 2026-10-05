// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
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
    private static readonly string[] DrawingPair = ["DrawFirst", "DrawSecond"];
    private static readonly string[] DrawingSelection = ["DrawFirst", "DrawSecond", "DrawThird"];
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
    private bool _disposed;

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

    [Fact]
    public async Task DaxMeasureWrites_NativeCommaSyntax_PreserveAndEvaluate()
    {
        var created = await CallToolAsync("file", new()
        {
            ["action"] = "create",
            ["path"] = _testExcelFile
        });
        AssertSuccess(created, "Create DAX workbook");
        var session = GetJsonProperty(created, "session_id");
        Assert.NotNull(session);
        var loaded = await CallToolAsync("powerquery", new()
        {
            ["action"] = "create",
            ["session_id"] = session,
            ["query_name"] = "SalesTable",
            ["m_code"] = "#table(type table [Amount = number], {{1000}, {2500}})",
            ["load_destination"] = "data-model"
        });
        AssertSuccess(loaded, "Load DAX source");
        const string formula = "DIVIDE(SUM(SalesTable[Amount]), 1000)";
        foreach (var update in new[] { false, true })
        {
            var name = $"Comma_{Guid.NewGuid():N}";
            var written = await CallToolAsync("datamodel", new()
            {
                ["action"] = "create-measure",
                ["session_id"] = session,
                ["table_name"] = "SalesTable",
                ["measure_name"] = name,
                ["dax_formula"] = update ? "SUM(SalesTable[Amount])" : formula,
                ["format_type"] = update ? "Percentage" : "Decimal"
            });
            AssertSuccess(written, "Create comma measure");
            if (update)
            {
                var updated = await CallToolAsync("datamodel", new()
                {
                    ["action"] = "update-measure",
                    ["session_id"] = session,
                    ["measure_name"] = name,
                    ["dax_formula"] = formula,
                    ["format_type"] = "Decimal"
                });
                AssertSuccess(updated, "Update comma measure");
            }
            var read = await CallToolAsync("datamodel_read", new()
            {
                ["action"] = "read",
                ["session_id"] = session,
                ["measure_name"] = name
            });
            AssertSuccess(read, "Read comma measure");
            Assert.Equal(formula, GetJsonProperty(read, "daxFormula"));
            using var readDocument = JsonDocument.Parse(read);
            Assert.Equal(name, readDocument.RootElement.GetProperty("measureName").GetString());
            Assert.Equal("SalesTable", readDocument.RootElement.GetProperty("tableName").GetString());
            Assert.Equal(formula.Length, readDocument.RootElement.GetProperty("characterCount").GetInt32());
            Assert.Equal("Decimal", readDocument.RootElement.GetProperty("formatInfo").GetProperty("type").GetString());
            var evaluated = await CallToolAsync("datamodel_read", new()
            {
                ["action"] = "evaluate",
                ["session_id"] = session,
                ["dax_query"] = $"EVALUATE ROW(\"Result\", [{name}])"
            });
            AssertSuccess(evaluated, "Evaluate comma measure");
            using var document = JsonDocument.Parse(evaluated);
            Assert.Equal(3.5m, Assert.Single(Assert.Single(
                document.RootElement.GetProperty("rows").EnumerateArray()).EnumerateArray()).GetDecimal());
        }
        var closed = await CallToolAsync("file", new()
        {
            ["action"] = "close",
            ["session_id"] = session,
            ["save"] = false
        });
        AssertSuccess(closed, "Close DAX workbook");
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

        // Flush telemetry before shutdown to ensure test telemetry is sent
        ExcelMcpTelemetry.Flush();
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
            ["path"] = _testExcelFile
        });
        AssertSuccess(createResult, "File creation and session open");
        Assert.True(File.Exists(_testExcelFile), "Excel file should exist");
        var sessionId = GetJsonProperty(createResult, "session_id");
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
            Assert.Equal(sessionId, session.GetProperty("session_id").GetString());
            Assert.False(session.TryGetProperty("sessionId", out _));
            Assert.True(session.GetProperty("canClose").GetBoolean());
        }
        await CallSuccessfulToolAsync("worksheet", new()
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data"
        });
        await CallSuccessfulToolAsync("range", new()
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1",
            ["values"] = new List<List<string>> { new() { "persisted-smoke-value" } }
        });
        await SaveAndVerifyAsync(sessionId, "persisted-smoke-value");
    }

    [Fact]
    [Trait("Acceptance", "Required")]
    public async Task Smoke_WorksheetAndRange_SearchAndOverwriteContracts()
    {
        var sessionId = await CreateSmokeWorkbookAsync();

        // =====================================================================
        // STEP 3: WORKSHEET OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 3: Worksheet operations...");

        var listSheetsResult = await CallToolAsync("worksheet_read", new Dictionary<string, object?>
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

        var findSheetResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["sheet_name"] = "FindSmoke"
        });
        AssertSuccess(findSheetResult, "Create find smoke worksheet");
        var findValuesResult = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "FindSmoke",
            ["range_address"] = "A1:A26",
            ["values"] = Enumerable.Range(0, 26)
                .Select(index => new[] { index < 25 ? "Apple" : "Banana" }).ToArray()
        });
        AssertSuccess(findValuesResult, "Write find smoke values");
        foreach (int? limit in new int?[] { null, 5 })
        {
            var findArguments = new Dictionary<string, object?>
            {
                ["action"] = "find",
                ["session_id"] = sessionId,
                ["sheet_name"] = "FindSmoke",
                ["range_address"] = "A1:A26",
                ["search_value"] = "Apple",
                ["find_options"] = new { matchEntireCell = true }
            };
            if (limit.HasValue)
            {
                findArguments["max_matches"] = limit.Value;
            }
            var findResult = await CallToolAsync("range_edit_read", findArguments);
            AssertSuccess(findResult, "Find bounded matches");
            using var findJson = JsonDocument.Parse(findResult);
            Assert.Equal(25, findJson.RootElement.GetProperty("totalCount").GetInt64());
            Assert.Equal(limit ?? 10, findJson.RootElement.GetProperty("returnedCount").GetInt32());
            Assert.Equal(limit ?? 10, findJson.RootElement.GetProperty("matchingCells").GetArrayLength());
            Assert.True(findJson.RootElement.GetProperty("truncated").GetBoolean());
        }
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
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C3",
            ["values"] = values
        });
        AssertSuccess(setValuesResult, "Set values");

        var rejectedWrite = await _client!.CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1",
            ["values"] = new List<List<string>> { new() { "Must not overwrite" } }
        }, cancellationToken: _cts.Token);
        Assert.True(rejectedWrite.IsError);
        var rejectionText = Assert.Single(rejectedWrite.Content.OfType<TextContentBlock>()).Text;
        using (var rejectionJson = JsonDocument.Parse(rejectionText))
        {
            Assert.False(rejectionJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("Conflict", rejectionJson.RootElement.GetProperty("errorCategory").GetString());
            Assert.Contains("$A$1", rejectionJson.RootElement.GetProperty("errorMessage").GetString());
        }
        var unchangedHeader = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1"
        });
        AssertSuccess(unchangedHeader, "Read header after rejected write");
        Assert.Equal("Name", GetFirstCellValue(unchangedHeader));
        var intentionalUpdate = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1",
            ["values"] = new List<List<string>> { new() { "Product" } },
            ["overwrite_policy"] = "allow"
        });
        AssertSuccess(intentionalUpdate, "Explicitly allowed header replacement");

        var getValuesResult = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C3"
        });
        AssertSuccess(getValuesResult, "Get values");
        Assert.Equal("Product", GetFirstCellValue(getValuesResult));
        using (var valuesJson = JsonDocument.Parse(getValuesResult))
        {
            var rows = valuesJson.RootElement.GetProperty("values");
            Assert.Equal(3, rows.GetArrayLength());
            Assert.Equal(3, rows[0].GetArrayLength());
            Assert.Equal("Item1", rows[1][0].GetString());
            Assert.Equal(100, rows[1][1].GetDouble());
            Assert.Equal("Item2", rows[2][0].GetString());
            Assert.Equal(200, rows[2][1].GetDouble());
        }
        _output.WriteLine("  ✓ range: SetValues and GetValues passed");
        await CloseSmokeWorkbookAsync(sessionId);
    }

    [Fact]
    [Trait("Acceptance", "Required")]
    public async Task Smoke_NativeFormattingAndStyles_RoundTrip()
    {
        var sessionId = await CreateSmokeWorkbookAsync(withData: true);
        var discoverCellsResult = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-special-cells",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C4",
            ["cell_kind"] = "constants"
        });
        AssertSuccess(discoverCellsResult, "Discover all constant cells");
        using (var discoveryJson = JsonDocument.Parse(discoverCellsResult))
        {
            Assert.Equal("Data", discoveryJson.RootElement.GetProperty("sheetName").GetString());
            Assert.Equal("$A$1:$C$4", discoveryJson.RootElement.GetProperty("rangeAddress").GetString());
            Assert.Equal("constants", discoveryJson.RootElement.GetProperty("cellKind").GetString());
            Assert.Equal(9, discoveryJson.RootElement.GetProperty("cellCount").GetInt64());
            var area = Assert.Single(discoveryJson.RootElement.GetProperty("areas").EnumerateArray());
            Assert.Equal("$A$1:$C$3", area.GetString());
        }
        _output.WriteLine("  ✓ range: SetValues and GetValues passed");
        var fineFormatResult = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_addresses"] = (string[])["AA10:AB11"],
            ["format_options"] = new
            {
                fontThemeColor = 5,
                fillThemeColor = 6,
                indentLevel = 2,
                horizontalAlignment = "left",
                borders = new[] { new { position = "DiagonalUp", lineStyle = "dash", color = "#123456" } }
            }
        });
        AssertSuccess(fineFormatResult, "Apply native fine formatting");
        var fineFormatRead = await CallToolAsync("range_format_read", new Dictionary<string, object?>
        {
            ["action"] = "get-format",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "AA10"
        });
        AssertSuccess(fineFormatRead, "Read native fine formatting");
        using (var fineJson = JsonDocument.Parse(fineFormatRead))
        {
            var cell = fineJson.RootElement.GetProperty("cells")[0].GetProperty("stored");
            Assert.Equal(5, cell.GetProperty("font").GetProperty("color").GetProperty("themeColor").GetInt32());
            Assert.Equal(6, cell.GetProperty("fill").GetProperty("color").GetProperty("themeColor").GetInt32());
            Assert.Equal(2, cell.GetProperty("indentLevel").GetInt32());
            var border = Assert.Single(cell.GetProperty("borders").EnumerateArray(),
                item => item.GetProperty("edge").GetString() == "xlDiagonalUp");
            Assert.Equal(-4115, border.GetProperty("lineStyle").GetInt32());
            Assert.Equal("#123456", border.GetProperty("color").GetProperty("rgb").GetString());
        }
        var styleCreate = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "create-cell-style",
            ["session_id"] = sessionId,
            ["style_name"] = "WorkflowHighlight",
            ["source_sheet_name"] = "Data",
            ["source_cell_address"] = "AA10"
        });
        AssertSuccess(styleCreate, "Create native cell style");
        var styleApply = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "set-style",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "AD10",
            ["style_name"] = "WorkflowHighlight"
        });
        AssertSuccess(styleApply, "Apply native custom style");
        var styleUpdate = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "update-cell-style",
            ["session_id"] = sessionId,
            ["style_name"] = "WorkflowHighlight",
            ["style_options"] = new { includeFont = true, formatOptions = new { bold = true } }
        });
        AssertSuccess(styleUpdate, "Update native custom style");
        var styleRead = await CallToolAsync("range_format_read", new Dictionary<string, object?>
        {
            ["action"] = "get-style",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "AD10"
        });
        AssertSuccess(styleRead, "Read native custom style user");
        using (var styleJson = JsonDocument.Parse(styleRead))
        {
            Assert.Equal("WorkflowHighlight", styleJson.RootElement.GetProperty("styleName").GetString());
            Assert.False(styleJson.RootElement.GetProperty("isBuiltInStyle").GetBoolean());
        }
        var styledCell = await CallSuccessfulToolAsync("range_format_read", new()
        {
            ["action"] = "get-format",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "AD10"
        });
        using (var cellJson = JsonDocument.Parse(styledCell))
            Assert.True(cellJson.RootElement.GetProperty("cells")[0].GetProperty("stored")
                .GetProperty("font").GetProperty("bold").GetBoolean());
        var styleDelete = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "delete-cell-style",
            ["session_id"] = sessionId,
            ["style_name"] = "WorkflowHighlight"
        });
        AssertSuccess(styleDelete, "Delete native custom style");
        var cellStyles = await CallSuccessfulToolAsync("workbook_read", new()
        {
            ["action"] = "list-cell-styles",
            ["session_id"] = sessionId
        });
        using (var stylesJson = JsonDocument.Parse(cellStyles))
            Assert.DoesNotContain(stylesJson.RootElement.GetProperty("styles").EnumerateArray(),
                style => style.GetProperty("name").GetString() == "WorkflowHighlight");
        AssertSuccess(await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "create-table-style",
            ["session_id"] = sessionId,
            ["style_name"] = "WorkflowTableStyle",
            ["source_style_name"] = "TableStyleMedium2"
        }), "Clone native table style");
        var tableStyleUpdate = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "update-table-style",
            ["session_id"] = sessionId,
            ["style_name"] = "WorkflowTableStyle",
            ["table_style_options"] = new { elements = new[] { new { elementType = "xlHeaderRow", fillColor = "#123456", bold = false } } }
        });
        AssertSuccess(tableStyleUpdate, "Update native table-style element");
        using (var definition = JsonDocument.Parse(tableStyleUpdate))
        {
            var elements = definition.RootElement.GetProperty("style").GetProperty("elements");
            Assert.Equal(43, elements.GetArrayLength());
            var header = Assert.Single(elements.EnumerateArray(), item => item.GetProperty("elementType").GetString() == "xlHeaderRow");
            Assert.Equal("#123456", header.GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
            Assert.False(header.GetProperty("font").GetProperty("bold").GetBoolean());
        }
        var tableStyleRead = await CallSuccessfulToolAsync("workbook_read", new Dictionary<string, object?>
        {
            ["action"] = "get-table-style",
            ["session_id"] = sessionId,
            ["style_name"] = "WorkflowTableStyle"
        });
        using (var definition = JsonDocument.Parse(tableStyleRead))
        {
            var style = definition.RootElement.GetProperty("style");
            Assert.Equal("WorkflowTableStyle", style.GetProperty("name").GetString());
            Assert.False(style.GetProperty("builtIn").GetBoolean());
            var header = Assert.Single(style.GetProperty("elements").EnumerateArray(),
                item => item.GetProperty("elementType").GetString() == "xlHeaderRow");
            Assert.Equal("#123456", header.GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
            Assert.False(header.GetProperty("font").GetProperty("bold").GetBoolean());
        }
        var tableStyles = await CallSuccessfulToolAsync("workbook_read", new Dictionary<string, object?>
        {
            ["action"] = "list-table-styles",
            ["session_id"] = sessionId
        });
        using (var catalogue = JsonDocument.Parse(tableStyles))
            Assert.Single(catalogue.RootElement.GetProperty("styles").EnumerateArray(),
                style => style.GetProperty("name").GetString() == "WorkflowTableStyle");
        AssertSuccess(await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "delete-table-style",
            ["session_id"] = sessionId,
            ["style_name"] = "WorkflowTableStyle"
        }), "Delete native custom table style");
        var afterDeletion = await CallSuccessfulToolAsync("workbook_read", new()
        {
            ["action"] = "list-table-styles",
            ["session_id"] = sessionId
        });
        using (var catalogue = JsonDocument.Parse(afterDeletion))
            Assert.DoesNotContain(catalogue.RootElement.GetProperty("styles").EnumerateArray(),
                style => style.GetProperty("name").GetString() == "WorkflowTableStyle");
        var formatReadResult = await CallToolAsync("range_format_read", new Dictionary<string, object?>
        {
            ["action"] = "get-format",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C3",
            ["view"] = "both"
        });
        AssertSuccess(formatReadResult, "Read complete stored and displayed formatting");
        using (var formatJson = JsonDocument.Parse(formatReadResult))
        {
            Assert.Equal(9, formatJson.RootElement.GetProperty("cellCount").GetInt64());
            var cells = formatJson.RootElement.GetProperty("cells");
            Assert.Equal(9, cells.GetArrayLength());
            Assert.True(cells[0].TryGetProperty("stored", out _));
            Assert.True(cells[0].TryGetProperty("displayed", out _));
        }
        await CloseSmokeWorkbookAsync(sessionId);
    }

    [Fact]
    [Trait("Acceptance", "Required")]
    public async Task Smoke_TablesNamedRangesAndReports_UseWorksheetData()
    {
        var sessionId = await CreateSmokeWorkbookAsync(withData: true);

        // =====================================================================
        // STEP 5: TABLE OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 5: Table operations...");
        var spillFormulaResult = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-formulas",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "G1",
            ["formulas"] = new List<List<string>> { new() { "=SEQUENCE(3)" } }
        });
        AssertSuccess(spillFormulaResult, "Write dynamic-array formula");
        var pasteTargetResult = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "I1",
            ["values"] = new List<List<int>> { new() { 123 } }
        });
        AssertSuccess(pasteTargetResult, "Prepare occupied formatting-only paste target");
        var pasteFormatsResult = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "copy",
            ["session_id"] = sessionId,
            ["source_sheet"] = "Data",
            ["source_range"] = "A1:C3",
            ["target_sheet"] = "Data",
            ["target_range"] = "I1",
            ["paste_kind"] = "formats"
        });
        AssertSuccess(pasteFormatsResult, "Native formatting-only paste");
        using (var pasteJson = JsonDocument.Parse(pasteFormatsResult))
        {
            Assert.Equal("$I$1:$K$3", pasteJson.RootElement.GetProperty("destinationAddress").GetString());
            Assert.Equal("formats", pasteJson.RootElement.GetProperty("pasteKind").GetString());
        }
        var pasteContentResult = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "I1"
        });
        AssertSuccess(pasteContentResult, "Read occupied paste target");
        using (var pasteContentJson = JsonDocument.Parse(pasteContentResult))
        {
            Assert.Equal(123d, pasteContentJson.RootElement.GetProperty("values")[0][0].GetDouble());
        }
        var spillReadResult = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-spill-info",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "G1:G3"
        });
        AssertSuccess(spillReadResult, "Read native spill source and result relationships");
        using (var spillJson = JsonDocument.Parse(spillReadResult))
        {
            Assert.Equal("supported", spillJson.RootElement.GetProperty("capability").GetString());
            var cells = spillJson.RootElement.GetProperty("cells");
            Assert.Equal(3, cells.GetArrayLength());
            Assert.Equal("source", cells[0].GetProperty("state").GetString());
            Assert.Equal("result", cells[2].GetProperty("state").GetString());
            Assert.Equal("$G$1", cells[2].GetProperty("sourceAddress").GetString());
            Assert.Equal("$G$1:$G$3", cells[2].GetProperty("spillAddress").GetString());
        }

        var patternSeeds = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "P1:P2",
            ["values"] = new List<List<int>> { new() { 1 }, new() { 3 } }
        });
        AssertSuccess(patternSeeds, "Write native pattern seeds");
        var patternFill = await CallToolAsync("range_edit", new Dictionary<string, object?>
        {
            ["action"] = "auto-fill",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["source_range"] = "P1:P2",
            ["destination_range"] = "P1:P4",
            ["fill_type"] = "series"
        });
        AssertSuccess(patternFill, "Extend native pattern");
        var patternRead = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "P4"
        });
        AssertSuccess(patternRead, "Read native pattern result");
        using (var patternJson = JsonDocument.Parse(patternRead))
        {
            Assert.Equal(7d, patternJson.RootElement.GetProperty("values")[0][0].GetDouble());
        }
        var relativeFormula = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-formulas",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "Q1",
            ["reference_style"] = "r1c1",
            ["formulas"] = new List<List<string>> { new() { "=RC[-1]*2" } }
        });
        AssertSuccess(relativeFormula, "Write native R1C1 formula");
        var directionalFill = await CallToolAsync("range_edit", new Dictionary<string, object?>
        {
            ["action"] = "fill",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "Q1:Q4",
            ["direction"] = "down"
        });
        AssertSuccess(directionalFill, "Fill relative formulas down");
        var formulaRead = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-formulas",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "Q4",
            ["reference_style"] = "r1c1"
        });
        AssertSuccess(formulaRead, "Read native R1C1 formula");
        using (var formulaJson = JsonDocument.Parse(formulaRead))
        {
            Assert.Equal("=RC[-1]*2", formulaJson.RootElement.GetProperty("formulas")[0][0].GetString());
            Assert.Equal(14, formulaJson.RootElement.GetProperty("values")[0][0].GetDouble());
        }
        var nativePrecedents = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "trace-precedents",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "Q4"
        });
        AssertSuccess(nativePrecedents, "Trace native local precedents");
        using (var traceJson = JsonDocument.Parse(nativePrecedents))
        {
            Assert.Equal(2, traceJson.RootElement.GetProperty("nodes").GetArrayLength());
            Assert.Single(traceJson.RootElement.GetProperty("edges").EnumerateArray());
            Assert.Empty(traceJson.RootElement.GetProperty("unresolved").EnumerateArray());
            var coverage = traceJson.RootElement.GetProperty("coverage");
            Assert.False(coverage.GetProperty("workbookComplete").GetBoolean());
            Assert.True(coverage.GetProperty("nativeTraversalComplete").GetBoolean());
        }
        var nativeDependents = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "trace-dependents",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "P4"
        });
        AssertSuccess(nativeDependents, "Trace native dependents with unresolved leaf coverage");
        using (var traceJson = JsonDocument.Parse(nativeDependents))
        {
            Assert.Equal(2, traceJson.RootElement.GetProperty("nodes").GetArrayLength());
            Assert.Single(traceJson.RootElement.GetProperty("edges").EnumerateArray());
            Assert.Single(traceJson.RootElement.GetProperty("unresolved").EnumerateArray());
            Assert.False(traceJson.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
        }

        var protectionFlags = await CallToolAsync("range_link", new Dictionary<string, object?>
        {
            ["action"] = "set-cell-protection",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "S1,S3",
            ["locked"] = false,
            ["formula_hidden"] = true
        });
        AssertSuccess(protectionFlags, "Set exact cell protection flags");
        var protectionRead = await CallToolAsync("range_link_read", new Dictionary<string, object?>
        {
            ["action"] = "get-cell-protection",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "S1:S3"
        });
        AssertSuccess(protectionRead, "Read complete cell protection flags");
        using (var protectionJson = JsonDocument.Parse(protectionRead))
        {
            var cells = protectionJson.RootElement.GetProperty("cells");
            Assert.Equal(3, cells.GetArrayLength());
            Assert.False(cells[0].GetProperty("locked").GetBoolean());
            Assert.True(cells[0].GetProperty("formulaHidden").GetBoolean());
            Assert.True(cells[1].GetProperty("locked").GetBoolean());
            Assert.False(cells[1].GetProperty("formulaHidden").GetBoolean());
        }
        try
        {
            var protectSheet = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
            {
                ["action"] = "set-protection",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Data",
                ["is_protected"] = true,
                ["options"] = new { allowFormattingRows = true }
            });
            AssertSuccess(protectSheet, "Protect sheet with row formatting permission");
            var readSheetProtection = await CallToolAsync("worksheet_style_read", new Dictionary<string, object?>
            {
                ["action"] = "get-protection",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Data"
            });
            AssertSuccess(readSheetProtection, "Read sheet permissions");
            using var permissionsJson = JsonDocument.Parse(readSheetProtection);
            Assert.True(permissionsJson.RootElement.GetProperty("protectContents").GetBoolean());
            Assert.True(permissionsJson.RootElement.GetProperty("permissions").GetProperty("allowFormattingRows").GetBoolean());
            Assert.False(permissionsJson.RootElement.GetProperty("permissions").GetProperty("allowSorting").GetBoolean());
        }
        finally
        {
            var unprotectSheet = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
            {
                ["action"] = "set-protection",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Data",
                ["is_protected"] = false
            });
            AssertSuccess(unprotectSheet, "Unprotect sheet for remaining workflow");
        }

        foreach (var (name, left) in new[] { ("DrawFirst", 20d), ("DrawSecond", 100d) })
        {
            var shape = await CallToolAsync("drawing", new Dictionary<string, object?>
            {
                ["action"] = "add-shape",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Data",
                ["name"] = name,
                ["left"] = left,
                ["width"] = 40d,
                ["height"] = 30d
            });
            AssertSuccess(shape, "Create drawing for layout");
        }
        var duplicatedDrawing = await CallToolAsync("drawing", new Dictionary<string, object?>
        {
            ["action"] = "duplicate-object",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["object_name"] = "DrawFirst",
            ["new_name"] = "DrawThird",
            ["offset_left"] = 180d
        });
        AssertSuccess(duplicatedDrawing, "Duplicate native drawing");
        var groupedDrawing = await CallToolAsync("drawing", new Dictionary<string, object?>
        {
            ["action"] = "group-objects",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["object_names"] = JsonSerializer.Serialize(DrawingPair),
            ["group_name"] = "DrawGroup"
        });
        AssertSuccess(groupedDrawing, "Group native drawings");
        using (var grouped = JsonDocument.Parse(groupedDrawing))
            Assert.Equal(2, grouped.RootElement.GetProperty("drawingObjects")[0].GetProperty("children").GetArrayLength());
        var ungroupedDrawing = await CallToolAsync("drawing", new Dictionary<string, object?>
        {
            ["action"] = "ungroup-object",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["object_name"] = "DrawGroup"
        });
        AssertSuccess(ungroupedDrawing, "Ungroup native drawings");
        var alignedDrawing = await CallToolAsync("drawing", new Dictionary<string, object?>
        {
            ["action"] = "align-objects",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["object_names"] = JsonSerializer.Serialize(DrawingSelection),
            ["alignment"] = "Top"
        });
        AssertSuccess(alignedDrawing, "Align selected drawings");
        var distributedDrawing = await CallToolAsync("drawing", new Dictionary<string, object?>
        {
            ["action"] = "distribute-objects",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["object_names"] = JsonSerializer.Serialize(DrawingSelection),
            ["distribution"] = "Horizontal"
        });
        AssertSuccess(distributedDrawing, "Distribute selected drawings");
        using (var distributed = JsonDocument.Parse(distributedDrawing))
            Assert.Equal(110d, distributed.RootElement.GetProperty("drawingObjects")[1].GetProperty("left").GetDouble(), 2);
        var orderedDrawing = await CallToolAsync("drawing", new Dictionary<string, object?>
        {
            ["action"] = "set-z-order",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["object_name"] = "DrawThird",
            ["z_order"] = "SendToBack"
        });
        AssertSuccess(orderedDrawing, "Reorder native drawing");
        using (var ordered = JsonDocument.Parse(orderedDrawing))
            Assert.Equal(1, ordered.RootElement.GetProperty("drawingObjects")[0].GetProperty("zOrderPosition").GetInt32());

        var nativeTheme = await CallToolAsync("workbook_read", new Dictionary<string, object?>
        {
            ["action"] = "get-theme",
            ["session_id"] = sessionId
        });
        AssertSuccess(nativeTheme, "Read complete native workbook theme");
        using (var themeJson = JsonDocument.Parse(nativeTheme))
        {
            Assert.Equal(12, themeJson.RootElement.GetProperty("colors").GetArrayLength());
            Assert.Equal(3, themeJson.RootElement.GetProperty("majorFonts").GetArrayLength());
            Assert.Equal(3, themeJson.RootElement.GetProperty("minorFonts").GetArrayLength());
        }
        var calculationSettings = await CallToolAsync("calculation_mode_read", new Dictionary<string, object?>
        {
            ["action"] = "get-settings",
            ["session_id"] = sessionId
        });
        AssertSuccess(calculationSettings, "Read native calculation settings");
        string previousMode;
        using (var settingsJson = JsonDocument.Parse(calculationSettings))
        {
            previousMode = settingsJson.RootElement.GetProperty("mode").GetString()!;
            Assert.Equal("application", settingsJson.RootElement.GetProperty("settingsScope").GetString());
        }
        try
        {
            var manualSettings = await CallToolAsync("calculation_mode", new Dictionary<string, object?>
            {
                ["action"] = "set-settings",
                ["session_id"] = sessionId,
                ["mode"] = "manual"
            });
            AssertSuccess(manualSettings, "Set manual calculation");
            var rebuildCalculation = await CallToolAsync("calculation_mode", new Dictionary<string, object?>
            {
                ["action"] = "calculate",
                ["session_id"] = sessionId,
                ["scope"] = "application",
                ["kind"] = "rebuild"
            });
            AssertSuccess(rebuildCalculation, "Rebuild native formula dependencies");
        }
        finally
        {
            var restoredSettings = await CallToolAsync("calculation_mode", new Dictionary<string, object?>
            {
                ["action"] = "set-settings",
                ["session_id"] = sessionId,
                ["mode"] = previousMode
            });
            AssertSuccess(restoredSettings, "Restore previous calculation mode");
        }
        var retainPrecision = await CallToolAsync("calculation_mode", new Dictionary<string, object?>
        {
            ["action"] = "set-precision",
            ["session_id"] = sessionId,
            ["precision_as_displayed"] = false
        });
        AssertSuccess(retainPrecision, "Keep stored numeric precision");

        var ownedContext = await CallToolAsync("window_read", new Dictionary<string, object?>
        {
            ["action"] = "get-context",
            ["session_id"] = sessionId
        });
        AssertSuccess(ownedContext, "Read owned window context");
        using (var contextJson = JsonDocument.Parse(ownedContext))
        {
            Assert.Equal("available", contextJson.RootElement.GetProperty("availability").GetString());
            Assert.NotEmpty(contextJson.RootElement.GetProperty("windows").EnumerateArray());
        }
        var hideRows = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "set-visibility",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A40,A42",
            ["axis"] = "rows",
            ["hidden"] = true
        });
        AssertSuccess(hideRows, "Hide exact rows");
        var inspectRows = await CallToolAsync("range_format_read", new Dictionary<string, object?>
        {
            ["action"] = "get-visibility",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A40:A42",
            ["axis"] = "rows"
        });
        AssertSuccess(inspectRows, "Read exact row visibility");
        using (var rowsJson = JsonDocument.Parse(inspectRows))
        {
            var rows = rowsJson.RootElement.GetProperty("items");
            Assert.Equal(3, rows.GetArrayLength());
            Assert.True(rows[0].GetProperty("hidden").GetBoolean());
            Assert.False(rows[1].GetProperty("hidden").GetBoolean());
            Assert.True(rows[2].GetProperty("hidden").GetBoolean());
            Assert.Equal("undetermined", rows[0].GetProperty("hiddenCause").GetString());
        }
        var showRows = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "set-visibility",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "A40,A42",
            ["axis"] = "rows",
            ["hidden"] = false
        });
        AssertSuccess(showRows, "Restore row visibility");

        var seriesSeed = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "R1",
            ["values"] = new List<List<int>> { new() { 5 } }
        });
        AssertSuccess(seriesSeed, "Write series seed");
        var seriesCreate = await CallToolAsync("range_edit", new Dictionary<string, object?>
        {
            ["action"] = "create-series",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "R1:R4",
            ["orientation"] = "columns",
            ["step_value"] = 5d
        });
        AssertSuccess(seriesCreate, "Create native DataSeries");
        var seriesRead = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "R4"
        });
        AssertSuccess(seriesRead, "Read native DataSeries");
        using (var seriesJson = JsonDocument.Parse(seriesRead))
        {
            Assert.Equal(20d, seriesJson.RootElement.GetProperty("values")[0][0].GetDouble());
        }

        var filterSheet = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Filters"
        });
        AssertSuccess(filterSheet, "Create filtering worksheet");
        var filterSeed = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Filters",
            ["range_address"] = "A1:B4",
            ["values"] = new List<List<object?>> { new() { "Category", "Amount" }, new() { "A", 10 }, new() { "B", 20 }, new() { "C", 30 } }
        });
        AssertSuccess(filterSeed, "Write filtering source");
        var filtered = await CallToolAsync("range_edit", new Dictionary<string, object?>
        {
            ["action"] = "apply-filter",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Filters",
            ["range_address"] = "A1:B4",
            ["column_index"] = 2,
            ["filter_options"] = new { filterOperator = "And", criteria1 = ">=20", criteria2 = "<=30" }
        });
        AssertSuccess(filtered, "Apply two native filter conditions");
        var filterRead = await CallToolAsync("range_edit_read", new Dictionary<string, object?>
        {
            ["action"] = "get-filters",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Filters",
            ["range_address"] = "A1:B4"
        });
        AssertSuccess(filterRead, "Read native filtering criteria");
        using (var filterJson = JsonDocument.Parse(filterRead))
        {
            var columns = filterJson.RootElement.GetProperty("columnFilters");
            Assert.Equal(2, columns.GetArrayLength());
            Assert.Equal(">=20", columns[1].GetProperty("criteria1").GetProperty("value").GetString());
            Assert.Equal("<=30", columns[1].GetProperty("criteria2").GetProperty("value").GetString());
        }

        var reportLayout = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "set-page-setup",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Filters",
            ["page_setup_options"] = new { printArea = "A1:B4", leftMargin = 36, centerHeader = "Report", zoomPercent = 100 }
        });
        AssertSuccess(reportLayout, "Set native report layout");
        var reportRead = await CallToolAsync("worksheet_style_read", new Dictionary<string, object?>
        {
            ["action"] = "get-page-setup",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Filters"
        });
        AssertSuccess(reportRead, "Read native report layout");
        using (var reportJson = JsonDocument.Parse(reportRead))
        {
            Assert.Equal("$A$1:$B$4", reportJson.RootElement.GetProperty("printArea").GetString());
            Assert.Equal(36d, reportJson.RootElement.GetProperty("leftMargin").GetDouble(), 2);
        }
        var duplicateSeed = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "X20:Y23",
            ["values"] = new List<List<object?>> { new() { "Key", "Amount" }, new() { 1, 10 }, new() { 1, 20 }, new() { 2, 30 } }
        });
        AssertSuccess(duplicateSeed, "Write duplicate records");
        var duplicates = await CallToolAsync("range_edit", new Dictionary<string, object?>
        {
            ["action"] = "remove-duplicates",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "X20:Y23",
            ["key_columns"] = new List<int> { 1 },
            ["has_headers"] = true
        });
        AssertSuccess(duplicates, "Remove duplicates");
        using (var duplicateJson = JsonDocument.Parse(duplicates))
        {
            Assert.Equal(1, duplicateJson.RootElement.GetProperty("removedRows").GetInt32());
            Assert.Equal(2, duplicateJson.RootElement.GetProperty("remainingRows").GetInt32());
            Assert.Equal("$X$20:$Y$22", duplicateJson.RootElement.GetProperty("remainingRange").GetString());
        }
        var parsingSeed = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "AB20:AB21",
            ["values"] = new List<List<string>> { new() { "001,10," }, new() { "002,20," } }
        });
        AssertSuccess(parsingSeed, "Write text parsing input");
        var parsed = await CallToolAsync("range_edit", new Dictionary<string, object?>
        {
            ["action"] = "text-to-columns",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["source_range"] = "AB20:AB21",
            ["destination_cell"] = "AD20",
            ["options"] = new { comma = true, fields = new[] { new { position = 1, dataType = "Text" } } }
        });
        AssertSuccess(parsed, "Split text");
        using (var parsedJson = JsonDocument.Parse(parsed))
        {
            Assert.Equal(3, parsedJson.RootElement.GetProperty("outputColumns").GetInt32());
            Assert.Equal("$AD$20:$AF$21", parsedJson.RootElement.GetProperty("destinationRange").GetString());
        }
        var parsedRead = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "AD20:AF21"
        });
        AssertSuccess(parsedRead, "Read parsed text");
        using (var parsedJson = JsonDocument.Parse(parsedRead))
        {
            var parsedValues = parsedJson.RootElement.GetProperty("values");
            Assert.Equal("001", parsedValues[0][0].GetString());
            Assert.Equal("002", parsedValues[1][0].GetString());
            Assert.Equal(20d, parsedValues[1][1].GetDouble());
        }

        var preflightTableResult = await CallToolAsync("table_read", new Dictionary<string, object?>
        {
            ["action"] = "preflight",
            ["session_id"] = sessionId,
            ["table_name"] = "DataTable",
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:B3",
            ["has_headers"] = true
        });
        AssertSuccess(preflightTableResult, "Preflight table");
        using (var preflightJson = JsonDocument.Parse(preflightTableResult))
        {
            Assert.True(preflightJson.RootElement.GetProperty("safeToCreate").GetBoolean());
            Assert.Equal("$A$1:$B$3", preflightJson.RootElement.GetProperty("effectiveRange").GetString());
            var finding = Assert.Single(
                preflightJson.RootElement.GetProperty("findings").EnumerateArray().ToArray());
            Assert.Equal("ExcludedContiguousColumns", finding.GetProperty("kind").GetString());
            Assert.Equal("Warning", finding.GetProperty("severity").GetString());
            Assert.True(finding.GetProperty("isHeuristic").GetBoolean());
            Assert.Equal(
                "$C$1:$C$3",
                Assert.Single(finding.GetProperty("addresses").EnumerateArray().ToArray()).GetString());
            Assert.False(string.IsNullOrWhiteSpace(finding.GetProperty("remediation").GetString()));
        }

        var createTableResult = await CallToolAsync("table", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["table_name"] = "DataTable",
            ["sheet_name"] = "Data",
            ["range_address"] = "A1:C3",
            ["has_headers"] = true
        });
        AssertSuccess(createTableResult, "Create table");

        var listTablesResult = await CallToolAsync("table_read", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listTablesResult, "List tables");
        _output.WriteLine("  ✓ table: Preflight, Create, and List passed");

        // =====================================================================
        // STEP 6: NAMED RANGE OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 6: Named range operations...");

        var createParamResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["name"] = "ReportDate",
            ["reference"] = "=Data!$C$2"
        });
        AssertSuccess(createParamResult, "Create named range");

        var readParamResult = await CallToolAsync("namedrange_read", new Dictionary<string, object?>
        {
            ["action"] = "read",
            ["session_id"] = sessionId,
            ["name"] = "ReportDate"
        });
        using (var namedRange = JsonDocument.Parse(readParamResult))
        {
            var root = namedRange.RootElement;
            Assert.Equal("ReportDate", root.GetProperty("name").GetString());
            Assert.Equal("=Data!$C$2", root.GetProperty("refersTo").GetString());
            Assert.Equal("Double", root.GetProperty("valueType").GetString());
            Assert.Equal(new DateTime(2024, 1, 1).ToOADate(), root.GetProperty("value").GetDouble());
        }
        _output.WriteLine("  ✓ namedrange: Create and Read passed");
        await VerifyReportsAsync(sessionId);
        await CloseSmokeWorkbookAsync(sessionId);
    }

    [Fact]
    [Trait("Acceptance", "Required")]
    public async Task Smoke_PowerQueryConnectionsAndModel_PreserveIdentity()
    {
        var sessionId = await CreateSmokeWorkbookAsync();

        // =====================================================================
        // STEP 7: POWER QUERY OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 7: Power Query operations...");

        // Create test CSV
        var csvContent = "Product,Quantity\nWidget,10\nGadget,20";
        await File.WriteAllTextAsync(_testCsvFile, csvContent);

        var mCode = $@"let
    Source = Csv.Document(File.Contents(""{_testCsvFile.Replace("\"", "\"\"")}""),[Delimiter="","", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    TypedColumns = Table.TransformColumnTypes(PromotedHeaders, {{ {{""Product"", type text}}, {{""Quantity"", Int64.Type}} }})
in
    TypedColumns";

        var createQueryResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["query_name"] = "CsvData",
            ["m_code"] = mCode,
            ["load_destination"] = "connection-only"
        });
        AssertSuccess(createQueryResult, "Create Power Query");

        var listQueriesResult = await CallToolAsync("powerquery_read", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listQueriesResult, "List Power Queries");
        using (var queries = JsonDocument.Parse(listQueriesResult))
        {
            var query = Assert.Single(queries.RootElement.GetProperty("queries").EnumerateArray());
            Assert.Equal("CsvData", query.GetProperty("name").GetString());
            Assert.True(query.GetProperty("isConnectionOnly").GetBoolean());
        }

        // Rename the query (US1: Power Query rename)
        var renameQueryResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = sessionId,
            ["old_name"] = "CsvData",
            ["new_name"] = "ProductData"
        });
        AssertSuccess(renameQueryResult, "Rename Power Query");

        // Verify rename by listing again
        var listAfterRenameResult = await CallToolAsync("powerquery_read", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listAfterRenameResult, "List Power Queries after rename");
        using (var queries = JsonDocument.Parse(listAfterRenameResult))
        {
            var query = Assert.Single(queries.RootElement.GetProperty("queries").EnumerateArray());
            Assert.Equal("ProductData", query.GetProperty("name").GetString());
            Assert.True(query.GetProperty("isConnectionOnly").GetBoolean());
        }
        var renamedCode = await CallSuccessfulToolAsync("powerquery_read", new()
        {
            ["action"] = "view",
            ["session_id"] = sessionId,
            ["query_name"] = "ProductData"
        });
        Assert.Equal(mCode.Replace("\r\n", "\n", StringComparison.Ordinal),
            GetJsonProperty(renamedCode, "mCode")!.Replace("\r\n", "\n", StringComparison.Ordinal));

        _output.WriteLine("  ✓ powerquery: Create, List, and Rename passed");

        // =====================================================================
        // STEP 8: CONNECTION OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 8: Connection operations...");

        var listConnectionsResult = await CallToolAsync("connection_read", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listConnectionsResult, "List connections");
        using (var connections = JsonDocument.Parse(listConnectionsResult))
        {
            Assert.Empty(connections.RootElement.GetProperty("connections").EnumerateArray());
        }
        _output.WriteLine("  ✓ connection: List passed");

        await VerifyModelAsync(sessionId);
        await CloseSmokeWorkbookAsync(sessionId);
    }

    private async Task VerifyReportsAsync(string sessionId)
    {
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

        AssertSuccess(await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "add-row-field",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Name"
        }), "Add PivotTable calculation row field");
        AssertSuccess(await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "add-value-field",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Value",
            ["custom_name"] = "Total Value"
        }), "Add named PivotTable calculation value field");
        var showValues = await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "set-field-calculation",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Total Value",
            ["calculation"] = "PercentOfTotal"
        });
        AssertSuccess(showValues, "Set native percentage of grand total");
        var pivotValues = await CallToolAsync("pivottable_calc_read", new Dictionary<string, object?>
        {
            ["action"] = "get-data",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot"
        });
        AssertSuccess(pivotValues, "Read actual native Pivot percentages");
        using (var pivotJson = JsonDocument.Parse(pivotValues))
        {
            var rows = pivotJson.RootElement.GetProperty("values");
            Assert.Equal(1d / 3d, rows[1][1].GetDouble(), precision: 8);
            Assert.Equal(2d / 3d, rows[2][1].GetDouble(), precision: 8);
        }
        var calculationFields = await CallToolAsync("pivottable_field_read", new Dictionary<string, object?>
        {
            ["action"] = "list-fields",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot"
        });
        AssertSuccess(calculationFields, "Read complete Values instances");
        var nativeLayout = await CallToolAsync("pivottable_calc", new Dictionary<string, object?>
        {
            ["action"] = "set-layout-options",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["layout_options"] = new { rowLayout = 1, repeatLabels = true, styleName = "PivotStyleMedium9", preserveFormatting = true }
        });
        AssertSuccess(nativeLayout, "Apply native Pivot style and repeated labels");
        using (var layoutJson = JsonDocument.Parse(nativeLayout))
            Assert.Equal("PivotStyleMedium9", layoutJson.RootElement.GetProperty("styleName").GetString());
        var nativeFilter = await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "add-field-filter",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Name",
            ["filter_options"] = new { type = "ValueIsGreaterThan", number1 = 15, dataFieldName = "Total Value" }
        });
        AssertSuccess(nativeFilter, "Add native Pivot value filter");
        using (var filterJson = JsonDocument.Parse(nativeFilter))
            Assert.Single(filterJson.RootElement.GetProperty("filters").EnumerateArray());
        AssertSuccess(await CallToolAsync("pivottable_field", new Dictionary<string, object?>
        {
            ["action"] = "clear-field-filters",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Name"
        }), "Clear selected native Pivot filters");
        var sourceReplacement = await CallToolAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "set-source",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["source_sheet_name"] = "Data",
            ["table_name"] = "DataTable"
        });
        AssertSuccess(sourceReplacement, "Replace only selected native Pivot cache");
        using (var sourceJson = JsonDocument.Parse(sourceReplacement))
            Assert.Equal("SalesPivot", Assert.Single(sourceJson.RootElement.GetProperty("sharedPivotTables").EnumerateArray()).GetString());
        AssertSuccess(await CallToolAsync("slicer", new Dictionary<string, object?>
        {
            ["action"] = "create-slicer",
            ["session_id"] = sessionId,
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Name",
            ["slicer_name"] = "ReportNames",
            ["destination_sheet"] = "Data",
            ["position"] = "AC1"
        }), "Create native report control");
        var controlUpdate = await CallToolAsync("slicer", new Dictionary<string, object?>
        {
            ["action"] = "update-slicer",
            ["session_id"] = sessionId,
            ["slicer_name"] = "ReportNames",
            ["slicer_options"] = new { width = 240, columnCount = 2, caption = "Names", displayHeader = false }
        });
        AssertSuccess(controlUpdate, "Update native control layout");
        var controlRead = await CallToolAsync("slicer_read", new Dictionary<string, object?>
        {
            ["action"] = "get-slicer",
            ["session_id"] = sessionId,
            ["slicer_name"] = "ReportNames"
        });
        AssertSuccess(controlRead, "Read complete native control");
        using (var controlJson = JsonDocument.Parse(controlRead))
        {
            var state = controlJson.RootElement.GetProperty("slicer");
            Assert.Equal("Names", state.GetProperty("caption").GetString());
            Assert.Equal(240d, state.GetProperty("width").GetDouble(), 2);
            Assert.Equal(2, state.GetProperty("columnCount").GetInt32());
        }
        using (var fieldJson = JsonDocument.Parse(calculationFields))
        {
            var value = Assert.Single(fieldJson.RootElement.GetProperty("valueFields").EnumerateArray());
            Assert.Equal("Total Value", value.GetProperty("fieldName").GetString());
            Assert.Equal("PercentOfTotal", value.GetProperty("calculation").GetString());
            Assert.Equal("Sum", value.GetProperty("function").GetString());
        }

        var listPivotsResult = await CallToolAsync("pivottable_read", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listPivotsResult, "List PivotTables");
        using (var pivots = JsonDocument.Parse(listPivotsResult))
        {
            var pivot = Assert.Single(pivots.RootElement.GetProperty("pivotTables").EnumerateArray());
            Assert.Equal("SalesPivot", pivot.GetProperty("name").GetString());
            Assert.Equal("Data", pivot.GetProperty("sheetName").GetString());
            Assert.Equal(1, pivot.GetProperty("rowFieldCount").GetInt32());
            Assert.Equal(1, pivot.GetProperty("valueFieldCount").GetInt32());
        }
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
        var createdChartRead = await CallSuccessfulToolAsync("chart_read", new()
        {
            ["action"] = "read",
            ["session_id"] = sessionId,
            ["chart_name"] = "DataChart"
        });
        using (var chart = JsonDocument.Parse(createdChartRead))
            Assert.Equal("ColumnClustered", chart.RootElement.GetProperty("chartType").GetString());

        var secondarySeries = await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-series-axis-group",
            ["session_id"] = sessionId,
            ["chart_name"] = "DataChart",
            ["series_index"] = 1,
            ["axis_group"] = "Secondary"
        });
        AssertSuccess(secondarySeries, "Assign native secondary axes");
        var formattedPoint = await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-point-format",
            ["session_id"] = sessionId,
            ["chart_name"] = "DataChart",
            ["series_index"] = 1,
            ["point_index"] = 1,
            ["point_options"] = new { fillColor = "#FF0000" }
        });
        AssertSuccess(formattedPoint, "Format selected chart point");
        using (var point = JsonDocument.Parse(formattedPoint))
            Assert.Equal("#FF0000", point.RootElement.GetProperty("fillColor").GetString());
        var errorBars = await CallToolAsync("chart_config", new Dictionary<string, object?>
        {
            ["action"] = "set-error-bars",
            ["session_id"] = sessionId,
            ["chart_name"] = "DataChart",
            ["series_index"] = 1,
            ["error_bar_options"] = new { kind = "Fixed", amount = 2d, endStyle = "NoCap" }
        });
        AssertSuccess(errorBars, "Set native chart error bars");
        using (var bars = JsonDocument.Parse(errorBars))
        {
            Assert.True(bars.RootElement.GetProperty("hasErrorBars").GetBoolean());
            Assert.False(bars.RootElement.GetProperty("settingsReadable").GetBoolean());
        }
        var chartImage = Path.Combine(_tempDir, "chart-image.png");
        var exportedChart = await CallToolAsync("chart", new Dictionary<string, object?>
        {
            ["action"] = "export-image",
            ["session_id"] = sessionId,
            ["chart_name"] = "DataChart",
            ["target_path"] = chartImage
        });
        AssertSuccess(exportedChart, "Export native chart image");
        Assert.True(new FileInfo(chartImage).Length > 1000);

        var listChartsResult = await CallToolAsync("chart_read", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        AssertSuccess(listChartsResult, "List Charts");
        using (var charts = JsonDocument.Parse(listChartsResult))
        {
            var chart = Assert.Single(charts.RootElement.GetProperty("charts").EnumerateArray());
            Assert.Equal("DataChart", chart.GetProperty("name").GetString());
            Assert.Equal("Data", chart.GetProperty("sheetName").GetString());
            Assert.False(chart.GetProperty("isPivotChart").GetBoolean());
            Assert.Equal(400d, chart.GetProperty("width").GetDouble(), 2);
            Assert.Equal(300d, chart.GetProperty("height").GetDouble(), 2);
        }
        var changedChart = await CallSuccessfulToolAsync("chart_read", new()
        {
            ["action"] = "read",
            ["session_id"] = sessionId,
            ["chart_name"] = "DataChart"
        });
        using (var chart = JsonDocument.Parse(changedChart))
        {
            var series = chart.RootElement.GetProperty("series");
            Assert.Equal(2, series.GetArrayLength());
            Assert.Equal("Secondary", series[0].GetProperty("axisGroup").GetString());
            Assert.Equal("Value", series[0].GetProperty("name").GetString());
            Assert.Equal(100d, series[0].GetProperty("values")[0].GetDouble());
            Assert.Equal(200d, series[0].GetProperty("values")[1].GetDouble());
        }
        _output.WriteLine("  ✓ chart: Create and List passed");

    }

    private async Task VerifyModelAsync(string sessionId)
    {
        // =====================================================================
        // STEP 11: DATA MODEL OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 11: Data Model operations...");

        var listDataModelResult = await CallToolAsync("datamodel_read", new Dictionary<string, object?>
        {
            ["action"] = "list-tables",
            ["session_id"] = sessionId
        });
        AssertSuccess(listDataModelResult, "List Data Model tables");
        using (var model = JsonDocument.Parse(listDataModelResult))
            Assert.Empty(model.RootElement.GetProperty("tables").EnumerateArray());

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
        var listAfterLoadResult = await CallToolAsync("datamodel_read", new Dictionary<string, object?>
        {
            ["action"] = "list-tables",
            ["session_id"] = sessionId
        });
        AssertSuccess(listAfterLoadResult, "List Data Model tables after load");
        using (var model = JsonDocument.Parse(listAfterLoadResult))
        {
            var table = Assert.Single(model.RootElement.GetProperty("tables").EnumerateArray());
            Assert.Equal("ProductData", table.GetProperty("name").GetString());
            Assert.Equal(2, table.GetProperty("recordCount").GetInt32());
        }
        var beforeRename = await ReadProductModelStateAsync(sessionId);

        var readModelConnectionResult = await CallToolAsync("datamodel_read", new Dictionary<string, object?>
        {
            ["action"] = "read-connection",
            ["session_id"] = sessionId
        });
        AssertSuccess(readModelConnectionResult, "Read Data Model connection");
        using (var connectionJson = JsonDocument.Parse(readModelConnectionResult))
        {
            Assert.Equal("MODEL", connectionJson.RootElement.GetProperty("connectionType").GetString());
            Assert.Equal(7, connectionJson.RootElement.GetProperty("connectionTypeValue").GetInt32());
            Assert.Contains(
                connectionJson.RootElement.GetProperty("tableNames").EnumerateArray(),
                table => table.GetString() == "ProductData");
        }

        // Attempt rename-table - this will return success=false due to Excel limitation
        var renameTableResult = await CallToolAsync("datamodel", new Dictionary<string, object?>
        {
            ["action"] = "rename-table",
            ["session_id"] = sessionId,
            ["old_name"] = "ProductData",
            ["new_name"] = "RenamedProductData"
        }, expectedError: true);
        // Expect JSON with success=false (not a crash)
        using var renameJson = JsonDocument.Parse(renameTableResult);
        Assert.True(renameJson.RootElement.TryGetProperty("success", out var renameSuccess));
        Assert.False(renameSuccess.GetBoolean(), "Rename-table should fail due to Excel limitation");
        Assert.True(renameJson.RootElement.TryGetProperty("errorMessage", out var renameError));
        var renameErrorText = renameError.GetString() ?? "";
        Assert.Contains("immutable", renameErrorText, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(beforeRename, await ReadProductModelStateAsync(sessionId));
        _output.WriteLine("  ✓ datamodel: RenameTable correctly returns error (Excel limitation)");

        _output.WriteLine("  ✓ datamodel: ListTables passed");
    }

    [Fact]
    [Trait("Acceptance", "Required")]
    public async Task Smoke_TypedFormattingAndInvalidMacroWorkbook_ReturnConcreteResults()
    {
        var sessionId = await CreateSmokeWorkbookAsync(withData: true);

        // =====================================================================
        // STEP 12: CONDITIONAL FORMAT OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 12: Conditional Format operations...");

        var addRuleResult = await CallToolAsync("conditionalformat", new Dictionary<string, object?>
        {
            ["action"] = "add-rule",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "B2:B3",
            ["rule_type"] = "cellvalue",
            ["operator_type"] = "greater",
            ["formula1"] = "100",
            ["interior_color"] = "#00FF00"
        });
        AssertSuccess(addRuleResult, "Add conditional format rule");
        _output.WriteLine("  ✓ conditionalformat: AddRule passed");

        var addTypedRuleResult = await CallToolAsync("conditionalformat", new Dictionary<string, object?>
        {
            ["action"] = "add-rule",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "C2:C3",
            ["rule_type"] = "top10",
            ["rank"] = 7,
            ["top10_percent"] = true,
            ["font_bold"] = true,
            ["font_italic"] = false
        });
        AssertSuccess(addTypedRuleResult, "Add typed conditional format rule");

        var listTypedRuleResult = await CallToolAsync("conditionalformat_read", new Dictionary<string, object?>
        {
            ["action"] = "list-rules",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Data",
            ["range_address"] = "C2:C3"
        });
        AssertSuccess(listTypedRuleResult, "List typed conditional format rule");

        using (var typedRuleJson = JsonDocument.Parse(listTypedRuleResult))
        {
            var typedRule = Assert.Single(typedRuleJson.RootElement.GetProperty("rules").EnumerateArray());
            Assert.Equal(7, typedRule.GetProperty("top10").GetProperty("rank").GetInt32());
            Assert.True(typedRule.GetProperty("top10").GetProperty("percent").GetBoolean());
            Assert.True(typedRule.GetProperty("fontBold").GetBoolean());
            Assert.False(typedRule.GetProperty("fontItalic").GetBoolean());
            Assert.Equal("$C$2:$C$3", typedRule.GetProperty("appliesTo").GetString());
            var updatedRuleResult = await CallToolAsync("conditionalformat", new Dictionary<string, object?>
            {
                ["action"] = "update-rule",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Data",
                ["rule_priority"] = typedRule.GetProperty("priority").GetInt32(),
                ["expected_fingerprint"] = typedRule.GetProperty("fingerprint").GetString(),
                ["options"] = new { rank = 5, stopIfTrue = false, appliesTo = "C2:C4" }
            });
            AssertSuccess(updatedRuleResult, "Update only selected conditional rule");
            using var updatedRules = JsonDocument.Parse(updatedRuleResult);
            var updated = Assert.Single(updatedRules.RootElement.GetProperty("rules").EnumerateArray(),
                rule => rule.GetProperty("type").GetString() == "top10");
            Assert.Equal(5, updated.GetProperty("top10").GetProperty("rank").GetInt32());
            Assert.False(updated.GetProperty("stopIfTrue").GetBoolean());
            Assert.Equal("$C$2:$C$4", updated.GetProperty("appliesTo").GetString());
            Assert.True(updated.GetProperty("fontBold").GetBoolean());
            Assert.False(updated.GetProperty("fontItalic").GetBoolean());
            var reorderedRuleResult = await CallToolAsync("conditionalformat", new Dictionary<string, object?>
            {
                ["action"] = "set-rule-priority",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Data",
                ["rule_priority"] = updated.GetProperty("priority").GetInt32(),
                ["expected_fingerprint"] = updated.GetProperty("fingerprint").GetString(),
                ["new_priority"] = 1
            });
            AssertSuccess(reorderedRuleResult, "Move selected conditional rule");
            using var reorderedRules = JsonDocument.Parse(reorderedRuleResult);
            var moved = Assert.Single(reorderedRules.RootElement.GetProperty("rules").EnumerateArray(),
                rule => rule.GetProperty("type").GetString() == "top10");
            Assert.Equal(1, moved.GetProperty("priority").GetInt32());
            Assert.Equal("$C$2:$C$4", moved.GetProperty("appliesTo").GetString());
            Assert.True(moved.GetProperty("fontBold").GetBoolean());
            Assert.False(moved.GetProperty("fontItalic").GetBoolean());
            var deletedRuleResult = await CallToolAsync("conditionalformat", new Dictionary<string, object?>
            {
                ["action"] = "delete-rule",
                ["session_id"] = sessionId,
                ["sheet_name"] = "Data",
                ["rule_priority"] = moved.GetProperty("priority").GetInt32(),
                ["expected_fingerprint"] = moved.GetProperty("fingerprint").GetString()
            });
            AssertSuccess(deletedRuleResult, "Delete only selected conditional rule");
            using var remainingRules = JsonDocument.Parse(deletedRuleResult);
            var remaining = Assert.Single(remainingRules.RootElement.GetProperty("rules").EnumerateArray());
            Assert.Equal("cellValue", remaining.GetProperty("type").GetString());
            Assert.Equal("=100", remaining.GetProperty("formula1").GetString());
            Assert.Equal("#00FF00", remaining.GetProperty("interiorColor").GetString());
        }
        _output.WriteLine("  ✓ conditionalformat: Typed integer/boolean arguments round-tripped");

        // =====================================================================
        // STEP 13: VBA OPERATIONS
        // =====================================================================
        _output.WriteLine("\n✓ Step 13: VBA operations...");

        var listVbaResult = await CallToolAsync("vba_read", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, expectedError: true);
        using (var listVbaJson = JsonDocument.Parse(listVbaResult))
        {
            Assert.False(listVbaJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(
                "InvalidInput",
                listVbaJson.RootElement.GetProperty("errorCategory").GetString());
            Assert.Contains(
                "macro-enabled",
                listVbaJson.RootElement.GetProperty("error").GetString());
        }
        _output.WriteLine("  ✓ vba: List rejected the unsupported .xlsx format");
        await CloseSmokeWorkbookAsync(sessionId);
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
        Assert.NotNull(verifySessionId);

        await VerifyAndCloseAsync(verifySessionId, async () =>
        {
            var finalSheetsResult = await CallToolAsync("worksheet_read", new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["session_id"] = verifySessionId
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
                ["session_id"] = verifySessionId,
                ["sheet_name"] = "Data",
                ["range_address"] = "A1"
            });
            Assert.Equal(expectedValue, GetFirstCellValue(values));
            _output.WriteLine("  Saved worksheet value survived reopening.");
        });

    }

    private async Task<string> CreateSmokeWorkbookAsync(bool withData = false)
    {
        var created = await CallSuccessfulToolAsync("file", new()
        {
            ["action"] = "create",
            ["path"] = _testExcelFile
        });
        var session = GetJsonProperty(created, "session_id");
        Assert.NotNull(session);
        if (withData)
        {
            await CallSuccessfulToolAsync("worksheet", new()
            {
                ["action"] = "create",
                ["session_id"] = session,
                ["sheet_name"] = "Data"
            });
            await CallSuccessfulToolAsync("range", new()
            {
                ["action"] = "set-values",
                ["session_id"] = session,
                ["sheet_name"] = "Data",
                ["range_address"] = "A1:C3",
                ["values"] = new object?[][]
                {
                    ["Name", "Value", "Date"],
                    ["Item1", 100, "2024-01-01"],
                    ["Item2", 200, "2024-01-02"]
                }
            });
        }
        return session;
    }

    private Task<string> CloseSmokeWorkbookAsync(string session) =>
        CallSuccessfulToolAsync("file", new()
        {
            ["action"] = "close",
            ["session_id"] = session,
            ["save"] = false
        });

    private async Task<string> CallSuccessfulToolAsync(string tool, Dictionary<string, object?> arguments)
    {
        var response = await CallToolAsync(tool, arguments);
        AssertSuccess(response, $"{tool}.{arguments["action"]}");
        return response;
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
        }, expectedError: true);

        _output.WriteLine($"Result: {result[..Math.Min(300, result.Length)]}...");

        // Should have success=false
        using var json = JsonDocument.Parse(result);
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

        // Step 1: Create source file with a sheet through the product entry point.
        _output.WriteLine("  1. Creating source file...");
        var createSource = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = sourceFile
        });
        AssertSuccess(createSource, "create source workbook");
        var sourceSessionId = GetJsonProperty(createSource, "session_id");
        Assert.NotNull(sourceSessionId);
        AssertSuccess(await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sourceSessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["values"] = new List<List<string>> { new() { "copied-content" } }
        }), "Write source worksheet content");
        AssertSuccess(await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = sourceSessionId,
            ["save"] = true
        }), "close source workbook");

        // Step 2: Create target file (empty, will receive the copied sheet)
        _output.WriteLine("  2. Creating target file...");
        var createTarget = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = targetFile
        });
        AssertSuccess(createTarget, "create target workbook");
        var targetSessionId = GetJsonProperty(createTarget, "session_id");
        Assert.NotNull(targetSessionId);
        AssertSuccess(await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = targetSessionId,
            ["save"] = false
        }), "close target workbook");

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
        var openedTarget = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = targetFile
        });
        AssertSuccess(openedTarget, "Open copied worksheet destination");
        var copiedSession = GetJsonProperty(openedTarget, "session_id");
        Assert.NotNull(copiedSession);
        var copiedValues = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = copiedSession,
            ["sheet_name"] = "CopiedSheet",
            ["range_address"] = "A1"
        });
        AssertSuccess(copiedValues, "Read copied worksheet contents");
        Assert.Equal("copied-content", GetFirstCellValue(copiedValues));
        AssertSuccess(await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = copiedSession,
            ["save"] = false
        }), "Close copied worksheet destination");
        _output.WriteLine("  ✓ copy-to-file succeeded WITHOUT session_id!");
        _output.WriteLine("✓ Atomic operation (copy-to-file) correctly works without session requirement!");
    }

    [Fact]
    public async Task VbaRun_OnMacroWorkbook_ViaMcpProtocol_ExecutesAndPersistsWorkbookSideEffect()
    {
        _output.WriteLine("\n✓ Testing vba.run end-to-end via MCP protocol...");

        var macroWorkbook = Path.Join(_tempDir, "VbaRunProof.xlsm");

        var createResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = macroWorkbook
        });
        AssertSuccess(createResult, "Create macro workbook");
        var sessionId = GetJsonProperty(createResult, "session_id");
        Assert.NotNull(sessionId);

        var importResult = await CallToolAsync("vba", new Dictionary<string, object?>
        {
            ["action"] = "import",
            ["session_id"] = sessionId,
            ["module_name"] = "TransportProof",
            ["vba_code"] = """
Sub WriteTransportProof()
    ThisWorkbook.Sheets(1).Range("A1").Value = "mcp-vba-run-ok"
End Sub
"""
        });
        AssertSuccess(importResult, "Import VBA module");

        var runResult = await CallToolAsync("vba", new Dictionary<string, object?>
        {
            ["action"] = "run",
            ["session_id"] = sessionId,
            ["procedure_name"] = "TransportProof.WriteTransportProof"
        });
        AssertSuccess(runResult, "Run VBA procedure");

        var getValuesResult = await CallToolAsync("range_read", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        });
        AssertSuccess(getValuesResult, "Read VBA side effect");
        Assert.Equal("mcp-vba-run-ok", GetFirstCellValue(getValuesResult));

        var closeResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = sessionId,
            ["save"] = true
        });
        AssertSuccess(closeResult, "Close macro workbook");

        var reopenedResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = macroWorkbook
        });
        AssertSuccess(reopenedResult, "Reopen macro workbook");
        var reopenedSessionId = GetJsonProperty(reopenedResult, "session_id");
        Assert.NotNull(reopenedSessionId);

        await VerifyAndCloseAsync(reopenedSessionId, async () =>
        {
            var persistedValueResult = await CallSuccessfulToolAsync("range_read", new Dictionary<string, object?>
            {
                ["action"] = "get-values",
                ["session_id"] = reopenedSessionId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1"
            });
            Assert.Equal("mcp-vba-run-ok", GetFirstCellValue(persistedValueResult));
        });
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

    private async Task<(string Code, string Queries, string Tables, string Connections)> ReadProductModelStateAsync(string sessionId)
    {
        var code = await CallSuccessfulToolAsync("powerquery_read", new()
        {
            ["action"] = "view",
            ["session_id"] = sessionId,
            ["query_name"] = "ProductData"
        });
        var queries = await CallSuccessfulToolAsync("powerquery_read", new()
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        var tables = await CallSuccessfulToolAsync("datamodel_read", new()
        {
            ["action"] = "list-tables",
            ["session_id"] = sessionId
        });
        var connections = await CallSuccessfulToolAsync("connection_read", new()
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        });
        var rows = await CallSuccessfulToolAsync("datamodel_read", new()
        {
            ["action"] = "evaluate",
            ["session_id"] = sessionId,
            ["dax_query"] = "EVALUATE SELECTCOLUMNS(ProductData, \"Product\", ProductData[Product], \"Quantity\", ProductData[Quantity]) ORDER BY [Product]"
        });
        using var queryJson = JsonDocument.Parse(queries);
        var query = Assert.Single(queryJson.RootElement.GetProperty("queries").EnumerateArray());
        Assert.Equal("ProductData", query.GetProperty("name").GetString());
        Assert.True(query.GetProperty("isLoadedToDataModel").GetBoolean());
        using var tableJson = JsonDocument.Parse(tables);
        var table = Assert.Single(tableJson.RootElement.GetProperty("tables").EnumerateArray());
        Assert.Equal("ProductData", table.GetProperty("name").GetString());
        Assert.Equal(2, table.GetProperty("recordCount").GetInt32());
        using var connectionJson = JsonDocument.Parse(connections);
        var identities = connectionJson.RootElement.GetProperty("connections").EnumerateArray()
            .Select(item => new
            {
                Name = item.GetProperty("name").GetString(),
                Type = item.GetProperty("type").GetString(),
                IsPowerQuery = item.GetProperty("isPowerQuery").GetBoolean()
            })
            .OrderBy(item => item.Name, StringComparer.Ordinal).ToArray();
        Assert.Equal("Query - ProductData", Assert.Single(identities, item => item.IsPowerQuery).Name);
        Assert.DoesNotContain(identities, item => item.Name?.Contains("RenamedProductData", StringComparison.Ordinal) == true);
        using var rowJson = JsonDocument.Parse(rows);
        var values = rowJson.RootElement.GetProperty("rows");
        Assert.Equal(2, values.GetArrayLength());
        Assert.Equal("Gadget", values[0][0].GetString());
        Assert.Equal(20, values[0][1].GetInt32());
        Assert.Equal("Widget", values[1][0].GetString());
        Assert.Equal(10, values[1][1].GetInt32());
        return (GetJsonProperty(code, "mCode")!, query.GetRawText(), table.GetRawText(),
            JsonSerializer.Serialize(identities));
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

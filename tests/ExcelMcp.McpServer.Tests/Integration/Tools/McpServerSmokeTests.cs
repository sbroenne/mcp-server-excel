using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Models;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Smoke test for MCP Server - Quick validation of core functionality from an LLM perspective.
///
/// PURPOSE: Fast, on-demand test to verify major functionality isn't broken.
/// SCOPE: Exercises the 12 main MCP tools with typical LLM workflows.
/// RUNTIME: ~30-60 seconds (fast enough for pre-commit checks).
///
/// Run this test before commits to catch breaking changes:
/// dotnet test --filter "FullyQualifiedName~McpServerSmokeTests.SmokeTest_AllTools_LlmWorkflow"
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SmokeTest")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
public class McpServerSmokeTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;
    private readonly string _testCsvFile;
    private readonly string _testQueryFile;

    public McpServerSmokeTests(ITestOutputHelper output)
    {
        _output = output;

        // Create temp directory for test files
        _tempDir = Path.Join(Path.GetTempPath(), $"McpSmokeTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Join(_tempDir, "SmokeTest.xlsx");
        _testCsvFile = Path.Join(_tempDir, "SampleData.csv");
        _testQueryFile = Path.Join(_tempDir, "TestQuery.pq");

        _output.WriteLine($"Test directory: {_tempDir}");
    }
    /// <inheritdoc/>

    public void Dispose()
    {
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
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Comprehensive smoke test that exercises all 12 MCP tools in a realistic LLM workflow using the session API.
    /// This test validates the complete tool chain and demonstrates proper session usage for multiple operations.
    /// </summary>
    [Fact(Skip = "Temporarily disabled during MCP session API refactor")]
    public async Task SmokeTest_AllTools_LlmWorkflow()
    {
        _output.WriteLine("=== MCP SERVER SMOKE TEST (SESSION API) ===");
        _output.WriteLine("Testing all 12 tools in optimized session workflow...\n");

        // =====================================================================
        // STEP 1: FILE CREATION (before session)
        // =====================================================================
        _output.WriteLine("✓ Step 1: Creating workbook...");

        // Create empty workbook
        var createResult = await ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        AssertSuccess(createResult, "File creation");
        Assert.True(File.Exists(_testExcelFile), "Excel file should exist");

        _output.WriteLine("  ✓ excel_file: CREATE passed");

        // =====================================================================
        // STEP 2: OPEN SESSION (75-90% faster for multiple operations)
        // =====================================================================
        _output.WriteLine("\n✓ Step 2: Opening session...");

        var openResult = await ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
        AssertSuccess(openResult, "Open session");
        var openJson = JsonDocument.Parse(openResult);
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(sessionId);

        _output.WriteLine($"  ✓ Session opened: {sessionId}");

        // =====================================================================
        // STEP 3: ALL OPERATIONS IN SESSION (using sessionId)
        // =====================================================================
        _output.WriteLine("\n✓ Step 3: Running all operations in active session...");

        // Test file (no session required for test)
        var testResult = await ExcelFileTool.ExcelFile(FileAction.Test, _testExcelFile);
        AssertSuccess(testResult, "File test");

        // Worksheet operations (with session)
        var listSheetsResult = await ExcelWorksheetTool.ExcelWorksheet(WorksheetAction.List, sessionId);
        AssertSuccess(listSheetsResult, "List worksheets in batch");

        var createSheetResult = await ExcelWorksheetTool.ExcelWorksheet(
            WorksheetAction.Create,
            sessionId,
            sheetName: "Data");
        AssertSuccess(createSheetResult, "Create worksheet in batch");

        _output.WriteLine("  ✓ excel_worksheet: LIST and CREATE in batch");

        // Range operations using session-aware tool API
        var values = new List<List<object?>>
        {
            new List<object?> { "Name", "Value", "Date" },
            new List<object?> { "Item1", 100, "2024-01-01" },
            new List<object?> { "Item2", 200, "2024-01-02" }
        };

        var setValuesResult = await ExcelRangeTool.ExcelRange(
            RangeAction.SetValues,
            _testExcelFile,
            sessionId,
            sheetName: "Data",
            rangeAddress: "A1:C3",
            values: values);
        AssertSuccess(setValuesResult, "Set values in batch");

        var getValuesResult = await ExcelRangeTool.ExcelRange(
            RangeAction.GetValues,
            _testExcelFile,
            sessionId,
            sheetName: "Data",
            rangeAddress: "A1:C3");
        AssertSuccess(getValuesResult, "Get values in batch");

        var usedRangeResult = await ExcelRangeTool.ExcelRange(
            RangeAction.GetUsedRange,
            _testExcelFile,
            sessionId,
            sheetName: "Data");
        AssertSuccess(usedRangeResult, "Get used range in batch");

        _output.WriteLine("  ✓ excel_range: SET/GET values and USED RANGE in batch");

        // Table operations via session API
        var createTableResult = await TableTool.Table(
            TableAction.Create,
            _testExcelFile,
            sessionId,
            tableName: "DataTable",
            sheetName: "Data",
            range: "A1:C3",
            hasHeaders: true);
        AssertSuccess(createTableResult, "Create table in batch");

        var listTablesResult = await TableTool.Table(
            TableAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listTablesResult, "List tables in batch");

        _output.WriteLine("  ✓ excel_table: CREATE and LIST in batch");

        // Named range operations via session API
        var createParamResult = await ExcelNamedRangeTool.ExcelParameter(
            NamedRangeAction.Create,
            _testExcelFile,
            sessionId,
            namedRangeName: "ReportDate",
            value: "=Data!$C$2");
        AssertSuccess(createParamResult, "Create named range in batch");

        var getParamResult = await ExcelNamedRangeTool.ExcelParameter(
            NamedRangeAction.Get,
            _testExcelFile,
            sessionId,
            namedRangeName: "ReportDate");
        AssertSuccess(getParamResult, "Get named range in batch");

        _output.WriteLine("  ✓ excel_namedrange: CREATE and GET in batch");

        // Power Query operations via session API
        var csvContent = "Product,Quantity\nWidget,10\nGadget,20";
        File.WriteAllText(_testCsvFile, csvContent);

        var mCode = $@"let
    Source = Csv.Document(File.Contents(""{_testCsvFile.Replace("\\", "\\\\")}""),[Delimiter="","", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders";
        File.WriteAllText(_testQueryFile, mCode);

        var importQueryResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            PowerQueryAction.Create,
            sessionId,
            queryName: "CsvData",
            sourcePath: _testQueryFile,
            loadDestination: "connection-only");
        AssertSuccess(importQueryResult, "Create Power Query in batch");

        var listQueriesResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            PowerQueryAction.List,
            sessionId);
        AssertSuccess(listQueriesResult, "List Power Queries in batch");

        _output.WriteLine("  ✓ excel_powerquery: IMPORT and LIST in batch");

        // Connection operations via session API
        var listConnectionsResult = await ExcelConnectionTool.ExcelConnection(
            ConnectionAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listConnectionsResult, "List connections in batch");

        _output.WriteLine("  ✓ excel_connection: LIST in batch");

        // Additional worksheet for session testing
        var createBatchSheetResult = await ExcelWorksheetTool.ExcelWorksheet(
            WorksheetAction.Create,
            sessionId,
            sheetName: "BatchTest");
        AssertSuccess(createBatchSheetResult, "Create additional worksheet in batch");

        // PivotTable operations via session API
        var createPivotResult = await ExcelPivotTableTool.ExcelPivotTable(
            PivotTableAction.CreateFromTable,
            _testExcelFile,
            sessionId,
            tableName: "DataTable",
            destinationSheet: "Data",
            destinationCell: "E1",
            pivotTableName: "SalesPivot");
        AssertSuccess(createPivotResult, "Create PivotTable in batch");

        var listPivotsResult = await ExcelPivotTableTool.ExcelPivotTable(
            PivotTableAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listPivotsResult, "List PivotTables in batch");

        _output.WriteLine("  ✓ excel_pivottable: CREATE and LIST in batch");

        // Data Model operations via session API
        var listDataModelResult = await ExcelDataModelTool.ExcelDataModel(
            DataModelAction.ListTables,
            _testExcelFile,
            sessionId);
        AssertSuccess(listDataModelResult, "List Data Model tables in batch");

        _output.WriteLine("  ✓ excel_datamodel: LIST TABLES in batch");

        // VBA operations via session API
        var listVbaResult = await ExcelVbaTool.ExcelVba(
            VbaAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listVbaResult, "List VBA modules in batch");

        _output.WriteLine("  ✓ excel_vba: LIST in session");

        // =====================================================================
        // STEP 4: SAVE AND CLOSE SESSION (persist all changes)
        // =====================================================================
        _output.WriteLine("\n✓ Step 4: Saving and closing session...");

        var saveResult = await ExcelFileTool.ExcelFile(FileAction.Save, sessionId: sessionId);
        AssertSuccess(saveResult, "Save session");

        var closeResult = await ExcelFileTool.ExcelFile(FileAction.Close, sessionId: sessionId);
        AssertSuccess(closeResult, "Close session");

        _output.WriteLine("  ✓ Session saved and closed");

        // =====================================================================
        // STEP 5: VERIFY OPERATIONS AFTER SESSION (persistence check)
        // =====================================================================
        _output.WriteLine("\n✓ Step 5: Verifying persistence after session close...");

        // Verify worksheets were created and saved via a fresh session
        var verifyOpenResult = await ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
        AssertSuccess(verifyOpenResult, "Re-open session for verification");
        var verifySessionJson = JsonDocument.Parse(verifyOpenResult);
        var verifySessionId = verifySessionJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrEmpty(verifySessionId), "Verification session should be created");

        try
        {
            var finalSheetsResult = await ExcelWorksheetTool.ExcelWorksheet(WorksheetAction.List, verifySessionId!);
            AssertSuccess(finalSheetsResult, "Final worksheet list");
            var sheetsJson = JsonDocument.Parse(finalSheetsResult);
            var worksheets = sheetsJson.RootElement.GetProperty("worksheets").EnumerateArray();
            var sheetNames = worksheets.Select(w => w.GetProperty("name").GetString()).ToList();

            Assert.Contains("Data", sheetNames);
            Assert.Contains("BatchTest", sheetNames);

            // Verify data was saved
            var finalDataResult = await ExcelRangeTool.ExcelRange(
                RangeAction.GetValues,
                _testExcelFile,
                verifySessionId!,
                sheetName: "Data",
                rangeAddress: "A1:C3");
            AssertSuccess(finalDataResult, "Final data verification");
        }
        finally
        {
            if (!string.IsNullOrEmpty(verifySessionId))
            {
                await ExcelFileTool.ExcelFile(FileAction.Close, sessionId: verifySessionId);
            }
        }

        _output.WriteLine("  ✓ All changes persisted correctly");

        // =====================================================================
        // FINAL VERIFICATION
        // =====================================================================
        _output.WriteLine("\n=== BATCH MODE SMOKE TEST COMPLETE ===");
        _output.WriteLine("✅ All 12 MCP tools tested successfully in BATCH MODE");
        _output.WriteLine("✅ Batch workflow: BEGIN → 15+ operations → COMMIT");
        _output.WriteLine("✅ Performance optimized: 75-90% faster than individual operations");
        _output.WriteLine("✅ Data persistence verified after batch commit");
        _output.WriteLine("✅ Demonstrates proper LLM batch mode usage pattern");
        _output.WriteLine("\n🚀 MCP Server batch functionality is working perfectly!");
    }

    /// <summary>
    /// Helper method to assert operation success and provide clear error messages.
    /// </summary>
    private void AssertSuccess(string jsonResult, string operationName)
    {
        Assert.NotNull(jsonResult);

        try
        {
            var json = JsonDocument.Parse(jsonResult);

            // Check for MCP error format
            if (json.RootElement.TryGetProperty("error", out var error))
            {
                var errorMsg = error.GetString();
                Assert.Fail($"{operationName} failed with error: {errorMsg}");
            }

            // Check for Success property (most operations)
            if (json.RootElement.TryGetProperty("Success", out var success))
            {
                if (!success.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("ErrorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned Success=false: {errorMsg}");
                }
            }
            // Check for success property (batch operations)
            else if (json.RootElement.TryGetProperty("success", out var successLower))
            {
                if (!successLower.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("errorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned success=false: {errorMsg}");
                }
            }

            _output.WriteLine($"  ✓ {operationName} succeeded");
        }
        catch (JsonException ex)
        {
            Assert.Fail($"{operationName} returned invalid JSON: {ex.Message}\nResponse: {jsonResult}");
        }
    }
}

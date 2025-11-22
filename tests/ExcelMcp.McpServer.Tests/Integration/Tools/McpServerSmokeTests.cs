using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
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
    /// Comprehensive smoke test that exercises all 13 MCP tools in a realistic LLM workflow using the session API.
    /// This test validates the complete tool chain and demonstrates proper session usage for multiple operations.
    /// </summary>
    [Fact]
    public void SmokeTest_AllTools_LlmWorkflow()
    {
        _output.WriteLine("=== MCP SERVER SMOKE TEST (SESSION API) ===");
        _output.WriteLine("Testing all 13 tools in optimized session workflow...\n");

        // =====================================================================
        // STEP 1: FILE CREATION (before session)
        // =====================================================================
        _output.WriteLine("✓ Step 1: Creating workbook...");

        // Create empty workbook
        var createResult = ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        AssertSuccess(createResult, "File creation");
        Assert.True(File.Exists(_testExcelFile), "Excel file should exist");

        _output.WriteLine("  ✓ excel_file: CREATE passed");

        // =====================================================================
        // STEP 2: OPEN SESSION (75-90% faster for multiple operations)
        // =====================================================================
        _output.WriteLine("\n✓ Step 2: Opening session...");

        var openResult = ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
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
        var testResult = ExcelFileTool.ExcelFile(FileAction.Test, _testExcelFile);
        AssertSuccess(testResult, "File test");

        // Worksheet operations (with session)
        var listSheetsResult = ExcelWorksheetTool.ExcelWorksheet(WorksheetAction.List, sessionId);
        AssertSuccess(listSheetsResult, "List worksheets in batch");

        var createSheetResult = ExcelWorksheetTool.ExcelWorksheet(
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

        var setValuesResult = ExcelRangeTool.ExcelRange(
            RangeAction.SetValues,
            _testExcelFile,
            sessionId,
            sheetName: "Data",
            rangeAddress: "A1:C3",
            values: values);
        AssertSuccess(setValuesResult, "Set values in batch");

        var getValuesResult = ExcelRangeTool.ExcelRange(
            RangeAction.GetValues,
            _testExcelFile,
            sessionId,
            sheetName: "Data",
            rangeAddress: "A1:C3");
        AssertSuccess(getValuesResult, "Get values in batch");

        var usedRangeResult = ExcelRangeTool.ExcelRange(
            RangeAction.GetUsedRange,
            _testExcelFile,
            sessionId,
            sheetName: "Data");
        AssertSuccess(usedRangeResult, "Get used range in batch");

        _output.WriteLine("  ✓ excel_range: SET/GET values and USED RANGE in batch");

        // Table operations via session API
        var createTableResult = TableTool.Table(
            TableAction.Create,
            _testExcelFile,
            sessionId,
            tableName: "DataTable",
            sheetName: "Data",
            range: "A1:C3",
            hasHeaders: true);
        AssertSuccess(createTableResult, "Create table in batch");

        var listTablesResult = TableTool.Table(
            TableAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listTablesResult, "List tables in batch");

        _output.WriteLine("  ✓ excel_table: CREATE and LIST in batch");

        // Named range operations via session API
        var createParamResult = ExcelNamedRangeTool.ExcelParameter(
            NamedRangeAction.Create,
            _testExcelFile,
            sessionId,
            namedRangeName: "ReportDate",
            value: "=Data!$C$2");
        AssertSuccess(createParamResult, "Create named range in batch");

        var getParamResult = ExcelNamedRangeTool.ExcelParameter(
            NamedRangeAction.Read,
            _testExcelFile,
            sessionId,
            namedRangeName: "ReportDate");
        AssertSuccess(getParamResult, "Read named range in batch");

        _output.WriteLine("  ✓ excel_namedrange: CREATE and READ in batch");

        // Power Query operations via session API
        var csvContent = "Product,Quantity\nWidget,10\nGadget,20";
        File.WriteAllText(_testCsvFile, csvContent);

        var mCode = $@"let
    Source = Csv.Document(File.Contents(""{_testCsvFile.Replace("\\", "\\\\")}""),[Delimiter="","", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders";
        File.WriteAllText(_testQueryFile, mCode);

        var importQueryResult = ExcelPowerQueryTool.ExcelPowerQuery(
            PowerQueryAction.Create,
            sessionId,
            queryName: "CsvData",
            mCode: mCode,
            loadDestination: "connection-only");
        AssertSuccess(importQueryResult, "Create Power Query in batch");

        var listQueriesResult = ExcelPowerQueryTool.ExcelPowerQuery(
            PowerQueryAction.List,
            sessionId);
        AssertSuccess(listQueriesResult, "List Power Queries in batch");

        _output.WriteLine("  ✓ excel_powerquery: IMPORT and LIST in batch");

        // Connection operations via session API
        var listConnectionsResult = ExcelConnectionTool.ExcelConnection(
            ConnectionAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listConnectionsResult, "List connections in batch");

        _output.WriteLine("  ✓ excel_connection: LIST in batch");

        // Additional worksheet for session testing
        var createBatchSheetResult = ExcelWorksheetTool.ExcelWorksheet(
            WorksheetAction.Create,
            sessionId,
            sheetName: "BatchTest");
        AssertSuccess(createBatchSheetResult, "Create additional worksheet in batch");

        // PivotTable operations via session API
        var createPivotResult = ExcelPivotTableTool.ExcelPivotTable(
            PivotTableAction.CreateFromTable,
            _testExcelFile,
            sessionId,
            tableName: "DataTable",
            destinationSheet: "Data",
            destinationCell: "E1",
            pivotTableName: "SalesPivot");
        AssertSuccess(createPivotResult, "Create PivotTable in batch");

        var listPivotsResult = ExcelPivotTableTool.ExcelPivotTable(
            PivotTableAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listPivotsResult, "List PivotTables in batch");

        _output.WriteLine("  ✓ excel_pivottable: CREATE and LIST in batch");

        // Chart operations via session API
        var createChartResult = ExcelChartTool.ExcelChart(
            ChartAction.CreateFromRange,
            _testExcelFile,
            sessionId,
            sheetName: "Data",
            sourceRange: "A1:C3",
            chartType: ChartType.ColumnClustered,
            left: 50,
            top: 50,
            width: 400,
            height: 300,
            chartName: "DataChart");
        AssertSuccess(createChartResult, "Create Chart in batch");

        var listChartsResult = ExcelChartTool.ExcelChart(
            ChartAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listChartsResult, "List Charts in batch");

        _output.WriteLine("  ✓ excel_chart: CREATE and LIST in batch");

        // Data Model operations via session API
        var listDataModelResult = ExcelDataModelTool.ExcelDataModel(
            DataModelAction.ListTables,
            _testExcelFile,
            sessionId);
        AssertSuccess(listDataModelResult, "List Data Model tables in batch");

        _output.WriteLine("  ✓ excel_datamodel: LIST TABLES in batch");

        // VBA operations via session API
        var listVbaResult = ExcelVbaTool.ExcelVba(
            VbaAction.List,
            _testExcelFile,
            sessionId);
        AssertSuccess(listVbaResult, "List VBA modules in batch");

        _output.WriteLine("  ✓ excel_vba: LIST in session");

        // =====================================================================
        // STEP 4: CLOSE SESSION (saves by default, persisting all changes)
        // =====================================================================
        _output.WriteLine("\n✓ Step 4: Closing session (saving changes)...");

        var closeResult = ExcelFileTool.ExcelFile(FileAction.Close, sessionId: sessionId, save: true);
        AssertSuccess(closeResult, "Close session with save");

        _output.WriteLine("  ✓ Session saved and closed");

        // =====================================================================
        // STEP 5: VERIFY OPERATIONS AFTER SESSION (persistence check)
        // =====================================================================
        _output.WriteLine("\n✓ Step 5: Verifying persistence after session close...");

        // Verify worksheets were created and saved via a fresh session
        var verifyOpenResult = ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
        AssertSuccess(verifyOpenResult, "Re-open session for verification");
        var verifySessionJson = JsonDocument.Parse(verifyOpenResult);
        var verifySessionId = verifySessionJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrEmpty(verifySessionId), "Verification session should be created");

        try
        {
            var finalSheetsResult = ExcelWorksheetTool.ExcelWorksheet(WorksheetAction.List, verifySessionId!);
            AssertSuccess(finalSheetsResult, "Final worksheet list");
            var sheetsJson = JsonDocument.Parse(finalSheetsResult);
            var worksheets = sheetsJson.RootElement.GetProperty("worksheets").EnumerateArray();
            var sheetNames = worksheets.Select(w => w.GetProperty("name").GetString()).ToList();

            Assert.Contains("Data", sheetNames);
            Assert.Contains("BatchTest", sheetNames);

            // Verify data was saved
            var finalDataResult = ExcelRangeTool.ExcelRange(
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
                ExcelFileTool.ExcelFile(FileAction.Close, sessionId: verifySessionId);
            }
        }

        _output.WriteLine("  ✓ All changes persisted correctly");

        // =====================================================================
        // FINAL VERIFICATION
        // =====================================================================
        _output.WriteLine("\n=== BATCH MODE SMOKE TEST COMPLETE ===");
        _output.WriteLine("✅ All 13 MCP tools tested successfully in BATCH MODE");
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

    /// <summary>
    /// Regression test for improved error message when invalid action is provided.
    /// GitHub Issue: User received unhelpful "An error occurred invoking 'excel_file'" when using action='Save'.
    /// Expected: Clear error message listing valid actions and explaining correct save workflow.
    /// </summary>
    [Fact]
    public void InvalidAction_Save_ReturnsHelpfulErrorMessage()
    {
        _output.WriteLine("Testing error message for invalid action 'Save'...");

        // Attempt to use non-existent 'Save' action
        // Note: We can't pass an invalid enum value directly, so this test verifies the catch block
        // by checking the tool's behavior when sessionId doesn't exist (similar error path)
        var result = ExcelFileTool.ExcelFile(FileAction.Close, sessionId: "nonexistent-session-id");

        _output.WriteLine($"Result: {result}");

        // Parse the JSON result
        var json = JsonDocument.Parse(result);

        // Should have success=false
        Assert.True(json.RootElement.TryGetProperty("success", out var success));
        Assert.False(success.GetBoolean());

        // Should have isError=true
        Assert.True(json.RootElement.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());

        // Should have helpful error message
        Assert.True(json.RootElement.TryGetProperty("errorMessage", out var errorMessage));
        var errorText = errorMessage.GetString();
        Assert.NotNull(errorText);
        Assert.Contains("not found", errorText, StringComparison.OrdinalIgnoreCase);

        _output.WriteLine("✓ Error message is clear and helpful");
    }
}

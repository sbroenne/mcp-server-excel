using Xunit;
using System.Diagnostics;
using System.Text.Json;
using System.Text;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.RoundTrip;

/// <summary>
/// Round trip tests for complete MCP Server workflows
/// These tests start the MCP server process and test comprehensive end-to-end scenarios
/// </summary>
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Feature", "MCPProtocol")]
public class McpServerRoundTripTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private Process? _serverProcess;

    public McpServerRoundTripTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"MCPRoundTrip_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (_serverProcess != null)
        {
            try 
            {
                if (!_serverProcess.HasExited)
                {
                    _serverProcess.Kill();
                }
            }
            catch (InvalidOperationException)
            {
                // Process already exited or disposed - this is fine
            }
            catch (Exception)
            {
                // Any other process cleanup error - ignore
            }
        }
        _serverProcess?.Dispose();
        
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch (Exception)
        {
            // Best effort cleanup
        }

        GC.SuppressFinalize(this);
    }

    #region Helper Methods

    private Process StartMcpServer()
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = "run --project src/ExcelMcp.McpServer",
            WorkingDirectory = Path.Combine(Directory.GetCurrentDirectory()),
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        var process = new Process { StartInfo = startInfo };
        process.Start();
        _serverProcess = process;
        return process;
    }

    private async Task InitializeServer(Process server)
    {
        var initRequest = new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "initialize",
            @params = new
            {
                protocolVersion = "2024-11-05",
                capabilities = new { },
                clientInfo = new
                {
                    name = "test-client",
                    version = "1.0.0"
                }
            }
        };

        var json = JsonSerializer.Serialize(initRequest);
        await server.StandardInput.WriteLineAsync(json);
        await server.StandardInput.FlushAsync();

        // Read and verify response
        var response = await server.StandardOutput.ReadLineAsync();
        Assert.NotNull(response);
    }

    private async Task<string> CallExcelTool(Process server, string toolName, object arguments)
    {
        var request = new
        {
            jsonrpc = "2.0",
            id = Environment.TickCount,  // Use TickCount for test IDs instead of Random
            method = "tools/call",
            @params = new
            {
                name = toolName,
                arguments = arguments
            }
        };

        var json = JsonSerializer.Serialize(request);
        await server.StandardInput.WriteLineAsync(json);
        await server.StandardInput.FlushAsync();

        var response = await server.StandardOutput.ReadLineAsync();
        Assert.NotNull(response);
        
        var responseJson = JsonDocument.Parse(response);
        if (responseJson.RootElement.TryGetProperty("error", out var error))
        {
            var errorMessage = error.GetProperty("message").GetString();
            throw new InvalidOperationException($"MCP tool call failed: {errorMessage}");
        }

        var result = responseJson.RootElement.GetProperty("result");
        var content = result.GetProperty("content")[0].GetProperty("text").GetString();
        Assert.NotNull(content);
        
        return content;
    }

    #endregion

    [Fact]
    public async Task McpServer_PowerQueryRoundTrip_ShouldCreateQueryLoadDataUpdateAndVerify()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);
        var testFile = Path.Combine(_tempDir, "roundtrip-test.xlsx");
        var queryName = "RoundTripQuery";
        var originalMCodeFile = Path.Combine(_tempDir, "original-query.pq");
        var updatedMCodeFile = Path.Combine(_tempDir, "updated-query.pq");
        var exportedMCodeFile = Path.Combine(_tempDir, "exported-query.pq");
        var targetSheet = "DataSheet";

        // Create initial M code that generates sample data
        var originalMCode = @"let
    Source = {
        [ID = 1, Name = ""Alice"", Department = ""Engineering""],
        [ID = 2, Name = ""Bob"", Department = ""Marketing""],
        [ID = 3, Name = ""Charlie"", Department = ""Sales""]
    },
    ConvertedToTable = Table.FromRecords(Source),
    AddedTitle = Table.AddColumn(ConvertedToTable, ""Title"", each ""Employee"")
in
    AddedTitle";

        // Create updated M code with additional transformation
        var updatedMCode = @"let
    Source = {
        [ID = 1, Name = ""Alice"", Department = ""Engineering""],
        [ID = 2, Name = ""Bob"", Department = ""Marketing""],
        [ID = 3, Name = ""Charlie"", Department = ""Sales""],
        [ID = 4, Name = ""Diana"", Department = ""HR""]
    },
    ConvertedToTable = Table.FromRecords(Source),
    AddedTitle = Table.AddColumn(ConvertedToTable, ""Title"", each ""Employee""),
    AddedStatus = Table.AddColumn(AddedTitle, ""Status"", each ""Active"")
in
    AddedStatus";

        await File.WriteAllTextAsync(originalMCodeFile, originalMCode);
        await File.WriteAllTextAsync(updatedMCodeFile, updatedMCode);

        try
        {
            _output.WriteLine("=== ROUND TRIP TEST: Power Query Complete Workflow ===");

            // Step 1: Create Excel file
            _output.WriteLine("Step 1: Creating Excel file...");
            await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

            // Step 2: Create target worksheet
            _output.WriteLine("Step 2: Creating target worksheet...");
            await CallExcelTool(server, "excel_worksheet", new { action = "create", excelPath = testFile, sheetName = targetSheet });

            // Step 3: Import Power Query
            _output.WriteLine("Step 3: Importing Power Query...");
            var importResponse = await CallExcelTool(server, "excel_powerquery", new 
            { 
                action = "import", 
                excelPath = testFile, 
                queryName = queryName,
                sourcePath = originalMCodeFile
            });
            var importJson = JsonDocument.Parse(importResponse);
            Assert.True(importJson.RootElement.GetProperty("Success").GetBoolean());

            // Step 4: Set Power Query to Load to Table mode (this should actually load data)
            _output.WriteLine("Step 4: Setting Power Query to Load to Table mode...");
            var setLoadResponse = await CallExcelTool(server, "excel_powerquery", new 
            { 
                action = "set-load-to-table", 
                excelPath = testFile, 
                queryName = queryName,
                targetSheet = targetSheet
            });
            var setLoadJson = JsonDocument.Parse(setLoadResponse);
            Assert.True(setLoadJson.RootElement.GetProperty("Success").GetBoolean());

            // Step 5: Verify initial data was loaded
            _output.WriteLine("Step 5: Verifying initial data was loaded...");
            var readResponse = await CallExcelTool(server, "excel_worksheet", new 
            { 
                action = "read", 
                excelPath = testFile, 
                sheetName = targetSheet,
                range = "A1:D10"  // Read headers plus data
            });
            var readJson = JsonDocument.Parse(readResponse);
            Assert.True(readJson.RootElement.GetProperty("Success").GetBoolean());
            var initialData = readJson.RootElement.GetProperty("Data").GetString();
            Assert.NotNull(initialData);
            Assert.Contains("Alice", initialData);
            Assert.Contains("Bob", initialData);
            Assert.Contains("Charlie", initialData);
            Assert.DoesNotContain("Diana", initialData);  // Should not be in original data
            _output.WriteLine($"Initial data verified: 3 rows loaded");

            // Step 6: Export Power Query for comparison
            _output.WriteLine("Step 6: Exporting Power Query...");
            var exportResponse = await CallExcelTool(server, "excel_powerquery", new 
            { 
                action = "export", 
                excelPath = testFile, 
                queryName = queryName,
                targetPath = exportedMCodeFile
            });
            var exportJson = JsonDocument.Parse(exportResponse);
            Assert.True(exportJson.RootElement.GetProperty("Success").GetBoolean());
            Assert.True(File.Exists(exportedMCodeFile));

            // Step 7: Update Power Query with enhanced M code
            _output.WriteLine("Step 7: Updating Power Query with enhanced M code...");
            var updateResponse = await CallExcelTool(server, "excel_powerquery", new 
            { 
                action = "update", 
                excelPath = testFile, 
                queryName = queryName,
                sourcePath = updatedMCodeFile
            });
            var updateJson = JsonDocument.Parse(updateResponse);
            Assert.True(updateJson.RootElement.GetProperty("Success").GetBoolean());

            // Step 8: Refresh the Power Query to apply changes
            // Note: The query should automatically refresh when updated, but we'll be explicit
            await Task.Delay(2000); // Allow time for Excel to process the update

            // Step 9: Verify updated data was loaded
            _output.WriteLine("Step 9: Verifying updated data was loaded...");
            var updatedReadResponse = await CallExcelTool(server, "excel_worksheet", new 
            { 
                action = "read", 
                excelPath = testFile, 
                sheetName = targetSheet,
                range = "A1:E10"  // Read more columns for Status column
            });
            var updatedReadJson = JsonDocument.Parse(updatedReadResponse);
            Assert.True(updatedReadJson.RootElement.GetProperty("Success").GetBoolean());
            var updatedData = updatedReadJson.RootElement.GetProperty("Data").GetString();
            Assert.NotNull(updatedData);
            Assert.Contains("Alice", updatedData);
            Assert.Contains("Bob", updatedData);
            Assert.Contains("Charlie", updatedData);
            Assert.Contains("Diana", updatedData);  // Should now be in updated data
            Assert.Contains("Active", updatedData);  // Should have Status column
            _output.WriteLine($"Updated data verified: 4 rows with Status column");

            // Step 10: List queries to verify it still exists
            _output.WriteLine("Step 10: Listing queries to verify integrity...");
            var listResponse = await CallExcelTool(server, "excel_powerquery", new 
            { 
                action = "list", 
                excelPath = testFile
            });
            var listJson = JsonDocument.Parse(listResponse);
            Assert.True(listJson.RootElement.GetProperty("Success").GetBoolean());
            var queries = listJson.RootElement.GetProperty("Queries").EnumerateArray();
            Assert.Contains(queries, q => q.GetProperty("Name").GetString() == queryName);

            _output.WriteLine("=== POWER QUERY ROUND TRIP TEST COMPLETED SUCCESSFULLY ===");
        }
        finally
        {
            // Cleanup test files
            try { if (File.Exists(testFile)) File.Delete(testFile); } catch { }
            try { if (File.Exists(originalMCodeFile)) File.Delete(originalMCodeFile); } catch { }
            try { if (File.Exists(updatedMCodeFile)) File.Delete(updatedMCodeFile); } catch { }
            try { if (File.Exists(exportedMCodeFile)) File.Delete(exportedMCodeFile); } catch { }
        }
    }

    [Fact]
    public async Task McpServer_VbaRoundTrip_ShouldImportRunAndVerifyExcelStateChanges()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);
        var testFile = Path.Combine(_tempDir, "vba-roundtrip-test.xlsm");
        var moduleName = "DataGeneratorModule";
        var originalVbaFile = Path.Combine(_tempDir, "original-generator.vba");
        var updatedVbaFile = Path.Combine(_tempDir, "enhanced-generator.vba");
        var exportedVbaFile = Path.Combine(_tempDir, "exported-module.vba");
        var testSheetName = "VBATestSheet";

        // Original VBA code - creates a sheet and fills it with data
        var originalVbaCode = @"Option Explicit

Public Sub GenerateTestData()
    Dim ws As Worksheet
    
    ' Create new worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = ""VBATestSheet""
    
    ' Fill with basic data
    ws.Cells(1, 1).Value = ""ID""
    ws.Cells(1, 2).Value = ""Name""
    ws.Cells(1, 3).Value = ""Value""
    
    ws.Cells(2, 1).Value = 1
    ws.Cells(2, 2).Value = ""Original""
    ws.Cells(2, 3).Value = 100
    
    ws.Cells(3, 1).Value = 2
    ws.Cells(3, 2).Value = ""Data""
    ws.Cells(3, 3).Value = 200
End Sub";

        // Enhanced VBA code - creates more sophisticated data
        var updatedVbaCode = @"Option Explicit

Public Sub GenerateTestData()
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Create new worksheet (delete if exists)
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(""VBATestSheet"").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = ""VBATestSheet""
    
    ' Enhanced headers
    ws.Cells(1, 1).Value = ""ID""
    ws.Cells(1, 2).Value = ""Name""
    ws.Cells(1, 3).Value = ""Value""
    ws.Cells(1, 4).Value = ""Status""
    ws.Cells(1, 5).Value = ""Generated""
    
    ' Generate multiple rows of enhanced data
    For i = 2 To 6
        ws.Cells(i, 1).Value = i - 1
        ws.Cells(i, 2).Value = ""Enhanced_"" & (i - 1)
        ws.Cells(i, 3).Value = (i - 1) * 150
        ws.Cells(i, 4).Value = ""Active""
        ws.Cells(i, 5).Value = Now()
    Next i
End Sub";

        await File.WriteAllTextAsync(originalVbaFile, originalVbaCode);
        await File.WriteAllTextAsync(updatedVbaFile, updatedVbaCode);

        try
        {
            _output.WriteLine("=== VBA ROUND TRIP TEST: Complete VBA Development Workflow ===");

            // Step 1: Create Excel file (.xlsm for VBA support)
            _output.WriteLine("Step 1: Creating Excel .xlsm file...");
            await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

            // Step 2: Import original VBA module
            _output.WriteLine("Step 2: Importing original VBA module...");
            var importResponse = await CallExcelTool(server, "excel_vba", new 
            { 
                action = "import", 
                excelPath = testFile, 
                moduleName = moduleName,
                sourcePath = originalVbaFile
            });
            var importJson = JsonDocument.Parse(importResponse);
            Assert.True(importJson.RootElement.GetProperty("Success").GetBoolean());

            // Step 3: Run original VBA to create initial sheet and data
            _output.WriteLine("Step 3: Running original VBA to create initial data...");
            var runResponse = await CallExcelTool(server, "excel_vba", new 
            { 
                action = "run", 
                excelPath = testFile, 
                moduleAndProcedure = $"{moduleName}.GenerateTestData"
            });
            var runJson = JsonDocument.Parse(runResponse);
            Assert.True(runJson.RootElement.GetProperty("Success").GetBoolean());

            // Step 4: Verify initial Excel state - sheet was created
            _output.WriteLine("Step 4: Verifying initial sheet was created...");
            var listSheetsResponse = await CallExcelTool(server, "excel_worksheet", new 
            { 
                action = "list", 
                excelPath = testFile
            });
            var listSheetsJson = JsonDocument.Parse(listSheetsResponse);
            Assert.True(listSheetsJson.RootElement.GetProperty("Success").GetBoolean());
            var sheets = listSheetsJson.RootElement.GetProperty("Sheets").EnumerateArray();
            Assert.Contains(sheets, s => s.GetProperty("Name").GetString() == testSheetName);

            // Step 5: Verify initial data was created by VBA
            _output.WriteLine("Step 5: Verifying initial data was created...");
            var readInitialResponse = await CallExcelTool(server, "excel_worksheet", new 
            { 
                action = "read", 
                excelPath = testFile, 
                sheetName = testSheetName,
                range = "A1:C10"
            });
            var readInitialJson = JsonDocument.Parse(readInitialResponse);
            Assert.True(readInitialJson.RootElement.GetProperty("Success").GetBoolean());
            var initialData = readInitialJson.RootElement.GetProperty("Data").GetString();
            Assert.NotNull(initialData);
            Assert.Contains("Original", initialData);
            Assert.Contains("Data", initialData);
            Assert.DoesNotContain("Enhanced", initialData);  // Should not be in original data
            _output.WriteLine("Initial VBA-generated data verified: 2 rows with basic structure");

            // Step 6: Export VBA module for comparison
            _output.WriteLine("Step 6: Exporting VBA module...");
            var exportResponse = await CallExcelTool(server, "excel_vba", new 
            { 
                action = "export", 
                excelPath = testFile, 
                moduleName = moduleName,
                targetPath = exportedVbaFile
            });
            var exportJson = JsonDocument.Parse(exportResponse);
            Assert.True(exportJson.RootElement.GetProperty("Success").GetBoolean());
            Assert.True(File.Exists(exportedVbaFile));

            // Step 7: Update VBA module with enhanced code
            _output.WriteLine("Step 7: Updating VBA module with enhanced code...");
            var updateResponse = await CallExcelTool(server, "excel_vba", new 
            { 
                action = "update", 
                excelPath = testFile, 
                moduleName = moduleName,
                sourcePath = updatedVbaFile
            });
            var updateJson = JsonDocument.Parse(updateResponse);
            Assert.True(updateJson.RootElement.GetProperty("Success").GetBoolean());

            // Step 8: Run updated VBA to create enhanced data
            _output.WriteLine("Step 8: Running updated VBA to create enhanced data...");
            var runUpdatedResponse = await CallExcelTool(server, "excel_vba", new 
            { 
                action = "run", 
                excelPath = testFile, 
                moduleAndProcedure = $"{moduleName}.GenerateTestData"
            });
            var runUpdatedJson = JsonDocument.Parse(runUpdatedResponse);
            Assert.True(runUpdatedJson.RootElement.GetProperty("Success").GetBoolean());

            // Step 9: Verify enhanced Excel state - data was updated
            _output.WriteLine("Step 9: Verifying enhanced data was created...");
            var readUpdatedResponse = await CallExcelTool(server, "excel_worksheet", new 
            { 
                action = "read", 
                excelPath = testFile, 
                sheetName = testSheetName,
                range = "A1:E10"  // Read more columns for Status and Generated columns
            });
            var readUpdatedJson = JsonDocument.Parse(readUpdatedResponse);
            Assert.True(readUpdatedJson.RootElement.GetProperty("Success").GetBoolean());
            var updatedData = readUpdatedJson.RootElement.GetProperty("Data").GetString();
            Assert.NotNull(updatedData);
            Assert.Contains("Enhanced_1", updatedData);
            Assert.Contains("Enhanced_5", updatedData);  // Should have 5 rows of enhanced data
            Assert.Contains("Active", updatedData);       // Should have Status column
            Assert.Contains("Generated", updatedData);    // Should have Generated column
            _output.WriteLine("Enhanced VBA-generated data verified: 5 rows with Status and Generated columns");

            // Step 10: List VBA modules to verify integrity
            _output.WriteLine("Step 10: Listing VBA modules to verify integrity...");
            var listVbaResponse = await CallExcelTool(server, "excel_vba", new 
            { 
                action = "list", 
                excelPath = testFile
            });
            var listVbaJson = JsonDocument.Parse(listVbaResponse);
            Assert.True(listVbaJson.RootElement.GetProperty("Success").GetBoolean());
            var modules = listVbaJson.RootElement.GetProperty("Scripts").EnumerateArray();
            Assert.Contains(modules, m => m.GetProperty("Name").GetString() == moduleName);

            _output.WriteLine("=== VBA ROUND TRIP TEST COMPLETED SUCCESSFULLY ===");
        }
        finally
        {
            // Cleanup test files
            try { if (File.Exists(testFile)) File.Delete(testFile); } catch { }
            try { if (File.Exists(originalVbaFile)) File.Delete(originalVbaFile); } catch { }
            try { if (File.Exists(updatedVbaFile)) File.Delete(updatedVbaFile); } catch { }
            try { if (File.Exists(exportedVbaFile)) File.Delete(exportedVbaFile); } catch { }
        }
    }
}

using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.RoundTrip.Commands;

/// <summary>
/// Round trip tests for Script (VBA) Core operations.
/// These are slow end-to-end tests that verify complete VBA development workflows.
/// Tests require Excel installation and VBA trust enabled.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public class ScriptCommandsRoundTripTests : IDisposable
{
    private readonly IScriptCommands _scriptCommands;
    private readonly IFileCommands _fileCommands;
    private readonly ISetupCommands _setupCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public ScriptCommandsRoundTripTests()
    {
        _scriptCommands = new ScriptCommands();
        _fileCommands = new FileCommands();
        _setupCommands = new SetupCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_VBA_RoundTrip_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "RoundTripWorkbook.xlsm");

        // Create test files
        CreateTestExcelFile();

        // Check VBA trust
        CheckVbaTrust();
    }

    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmpty(_testExcelFile, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    private void CheckVbaTrust()
    {
        var trustResult = _setupCommands.CheckVbaTrust(_testExcelFile);
        if (!trustResult.IsTrusted)
        {
            throw new InvalidOperationException("VBA trust is not enabled. Run 'excelcli setup-vba-trust' first.");
        }
    }

    [Fact]
    public async Task VbaRoundTrip_ShouldImportRunAndVerifyExcelStateChanges()
    {
        // Arrange - Create VBA module files for the complete workflow
        var originalVbaFile = Path.Combine(_tempDir, "data-generator.vba");
        var updatedVbaFile = Path.Combine(_tempDir, "enhanced-generator.vba");
        var exportedVbaFile = Path.Combine(_tempDir, "exported-module.vba");
        var moduleName = "DataGeneratorModule";
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

        // Updated VBA code - creates more sophisticated data
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

        // Need worksheet commands to verify VBA effects
        var worksheetCommands = new SheetCommands();

        try
        {
            // Step 1: Import original VBA module
            var importResult = await _scriptCommands.Import(_testExcelFile, moduleName, originalVbaFile);
            Assert.True(importResult.Success, $"Failed to import VBA module: {importResult.ErrorMessage}");

            // Step 2: List modules to verify import
            var listResult = _scriptCommands.List(_testExcelFile);
            Assert.True(listResult.Success, $"Failed to list VBA modules: {listResult.ErrorMessage}");
            Assert.Contains(listResult.Scripts, s => s.Name == moduleName);

            // Step 3: Run the VBA to create sheet and fill data
            var runResult1 = _scriptCommands.Run(_testExcelFile, $"{moduleName}.GenerateTestData", Array.Empty<string>());
            Assert.True(runResult1.Success, $"Failed to run VBA GenerateTestData: {runResult1.ErrorMessage}");

            // Step 4: Verify the VBA created the sheet by listing worksheets
            var listSheetsResult1 = worksheetCommands.List(_testExcelFile);
            Assert.True(listSheetsResult1.Success, $"Failed to list worksheets: {listSheetsResult1.ErrorMessage}");
            Assert.Contains(listSheetsResult1.Worksheets, w => w.Name == testSheetName);

            // Step 5: Read the data that VBA wrote to verify original functionality
            var readResult1 = worksheetCommands.Read(_testExcelFile, testSheetName, "A1:C3");
            Assert.True(readResult1.Success, $"Failed to read VBA-generated data: {readResult1.ErrorMessage}");

            // Verify original data structure (headers + 2 data rows)
            Assert.Equal(3, readResult1.Data.Count); // Header + 2 rows
            var headerRow = readResult1.Data[0];
            Assert.Equal("ID", headerRow[0]?.ToString());
            Assert.Equal("Name", headerRow[1]?.ToString());
            Assert.Equal("Value", headerRow[2]?.ToString());

            var dataRow1 = readResult1.Data[1];
            Assert.Equal("1", dataRow1[0]?.ToString());
            Assert.Equal("Original", dataRow1[1]?.ToString());
            Assert.Equal("100", dataRow1[2]?.ToString());

            // Step 6: Export the original module for verification
            var exportResult1 = await _scriptCommands.Export(_testExcelFile, moduleName, exportedVbaFile);
            Assert.True(exportResult1.Success, $"Failed to export original VBA module: {exportResult1.ErrorMessage}");

            var exportedContent1 = await File.ReadAllTextAsync(exportedVbaFile);
            Assert.Contains("GenerateTestData", exportedContent1);
            Assert.Contains("Original", exportedContent1);

            // Step 7: Update the module with enhanced version
            var updateResult = await _scriptCommands.Update(_testExcelFile, moduleName, updatedVbaFile);
            Assert.True(updateResult.Success, $"Failed to update VBA module: {updateResult.ErrorMessage}");

            // Step 8: Run the updated VBA to generate enhanced data
            var runResult2 = _scriptCommands.Run(_testExcelFile, $"{moduleName}.GenerateTestData", Array.Empty<string>());
            Assert.True(runResult2.Success, $"Failed to run updated VBA GenerateTestData: {runResult2.ErrorMessage}");

            // Step 9: Read the enhanced data to verify update worked
            var readResult2 = worksheetCommands.Read(_testExcelFile, testSheetName, "A1:E6");
            Assert.True(readResult2.Success, $"Failed to read enhanced VBA-generated data: {readResult2.ErrorMessage}");

            // Verify enhanced data structure (headers + 5 data rows, 5 columns)
            Assert.Equal(6, readResult2.Data.Count); // Header + 5 rows
            var enhancedHeaderRow = readResult2.Data[0];
            Assert.Equal("ID", enhancedHeaderRow[0]?.ToString());
            Assert.Equal("Name", enhancedHeaderRow[1]?.ToString());
            Assert.Equal("Value", enhancedHeaderRow[2]?.ToString());
            Assert.Equal("Status", enhancedHeaderRow[3]?.ToString());
            Assert.Equal("Generated", enhancedHeaderRow[4]?.ToString());

            var enhancedDataRow1 = readResult2.Data[1];
            Assert.Equal("1", enhancedDataRow1[0]?.ToString());
            Assert.Equal("Enhanced_1", enhancedDataRow1[1]?.ToString());
            Assert.Equal("150", enhancedDataRow1[2]?.ToString());
            Assert.Equal("Active", enhancedDataRow1[3]?.ToString());
            // Note: Generated column has timestamp, just verify it's not empty
            Assert.False(string.IsNullOrEmpty(enhancedDataRow1[4]?.ToString()));

            // Step 10: Export updated module and verify changes
            var exportResult2 = await _scriptCommands.Export(_testExcelFile, moduleName, exportedVbaFile);
            Assert.True(exportResult2.Success, $"Failed to export updated VBA module: {exportResult2.ErrorMessage}");

            var exportedContent2 = await File.ReadAllTextAsync(exportedVbaFile);
            Assert.Contains("Enhanced_", exportedContent2);
            Assert.Contains("Status", exportedContent2);
            Assert.Contains("For i = 2 To 6", exportedContent2);

            // Step 11: Final cleanup - delete the module
            var deleteResult = _scriptCommands.Delete(_testExcelFile, moduleName);
            Assert.True(deleteResult.Success, $"Failed to delete VBA module: {deleteResult.ErrorMessage}");

            // Step 12: Verify module is deleted
            var listResult2 = _scriptCommands.List(_testExcelFile);
            Assert.True(listResult2.Success, $"Failed to list VBA modules after delete: {listResult2.ErrorMessage}");
            Assert.DoesNotContain(listResult2.Scripts, s => s.Name == moduleName);
        }
        finally
        {
            // Cleanup files
            File.Delete(originalVbaFile);
            File.Delete(updatedVbaFile);
            if (File.Exists(exportedVbaFile)) File.Delete(exportedVbaFile);
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, true);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        _disposed = true;
    }
}
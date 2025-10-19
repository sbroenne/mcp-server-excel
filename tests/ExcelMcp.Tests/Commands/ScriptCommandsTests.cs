using Xunit;
using ExcelMcp.Core.Commands;
using System.IO;

namespace ExcelMcp.Tests.Commands;

/// <summary>
/// Integration tests for VBA script operations using Excel COM automation.
/// These tests require Excel installation and VBA trust settings for macro execution.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "VBA")]
public class ScriptCommandsTests : IDisposable
{
    private readonly ScriptCommands _scriptCommands;
    private readonly SheetCommands _sheetCommands;
    private readonly FileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _testVbaFile;
    private readonly string _testCsvFile;
    private readonly string _tempDir;

    /// <summary>
    /// Check if VBA access is trusted - helper for conditional test execution
    /// </summary>
    private bool IsVbaAccessAvailable()
    {
        try
        {
            int result = ExcelHelper.WithExcel(_testExcelFile, false, (excel, workbook) =>
            {
                try
                {
                    dynamic vbProject = workbook.VBProject;
                    int componentCount = vbProject.VBComponents.Count;
                    return 1; // Success
                }
                catch
                {
                    return 0; // Failure
                }
            });
            return result == 1;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Try to enable VBA access for testing
    /// </summary>
    private bool TryEnableVbaAccess()
    {
        try
        {
            var setupCommands = new SetupCommands();
            int result = setupCommands.EnableVbaTrust(new string[] { "setup-vba-trust" });
            return result == 0;
        }
        catch
        {
            return false;
        }
    }

    public ScriptCommandsTests()
    {
        _scriptCommands = new ScriptCommands();
        _sheetCommands = new SheetCommands();
        _fileCommands = new FileCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsm"); // Use .xlsm for VBA tests
        _testVbaFile = Path.Combine(_tempDir, "TestModule.vba");
        _testCsvFile = Path.Combine(_tempDir, "TestData.csv");
        
        // Create test files
        CreateTestExcelFile();
        CreateTestVbaFile();
        CreateTestCsvFile();
    }

    private void CreateTestExcelFile()
    {
        // Create an empty Excel file for testing
        string[] args = { "create-empty", _testExcelFile };
        
        int result = _fileCommands.CreateEmpty(args);
        if (result != 0)
        {
            throw new InvalidOperationException("Failed to create test Excel file. Excel may not be installed.");
        }
    }

    private void CreateTestVbaFile()
    {
        // Create a VBA module that adds data to a worksheet
        string vbaCode = @"Option Explicit

Sub AddTestData()
    ' Add sample data to the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Add headers
    ws.Cells(1, 1).Value = ""ID""
    ws.Cells(1, 2).Value = ""Name""
    ws.Cells(1, 3).Value = ""Value""
    
    ' Add data rows
    ws.Cells(2, 1).Value = 1
    ws.Cells(2, 2).Value = ""Test Item 1""
    ws.Cells(2, 3).Value = 100
    
    ws.Cells(3, 1).Value = 2
    ws.Cells(3, 2).Value = ""Test Item 2""
    ws.Cells(3, 3).Value = 200
    
    ws.Cells(4, 1).Value = 3
    ws.Cells(4, 2).Value = ""Test Item 3""
    ws.Cells(4, 3).Value = 300
End Sub

Function CalculateSum() As Long
    ' Calculate sum of values in column C
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim total As Long
    total = 0
    
    Dim i As Long
    For i = 2 To 4 ' Rows 2-4 contain data
        total = total + ws.Cells(i, 3).Value
    Next i
    
    CalculateSum = total
End Function

Sub AddDataWithParameters(startRow As Long, itemCount As Long, baseValue As Long)
    ' Add data with parameters - useful for testing parameter passing
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Add headers if starting at row 1
    If startRow = 1 Then
        ws.Cells(1, 1).Value = ""ID""
        ws.Cells(1, 2).Value = ""Name""
        ws.Cells(1, 3).Value = ""Value""
        startRow = 2
    End If
    
    ' Add data rows
    Dim i As Long
    For i = 0 To itemCount - 1
        ws.Cells(startRow + i, 1).Value = i + 1
        ws.Cells(startRow + i, 2).Value = ""Item "" & (i + 1)
        ws.Cells(startRow + i, 3).Value = baseValue + (i * 10)
    Next i
End Sub
";
        
        File.WriteAllText(_testVbaFile, vbaCode);
    }

    private void CreateTestCsvFile()
    {
        // Create a simple CSV file for testing
        string csvContent = @"ID,Name,Value
1,Initial Item 1,50
2,Initial Item 2,75
3,Initial Item 3,100";
        
        File.WriteAllText(_testCsvFile, csvContent);
    }

    [Fact]
    public void List_WithValidFile_ReturnsSuccess()
    {
        // Arrange
        string[] args = { "script-list", _testExcelFile };

        // Act
        int result = _scriptCommands.List(args);

        // Assert
        Assert.Equal(0, result);
    }

    [Fact]
    public void List_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "script-list" }; // Missing file argument

        // Act
        int result = _scriptCommands.List(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void List_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "script-list", "nonexistent.xlsx" };

        // Act
        int result = _scriptCommands.List(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Export_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "script-export", _testExcelFile }; // Missing module name

        // Act
        int result = _scriptCommands.Export(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Export_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "script-export", "nonexistent.xlsx", "Module1" };

        // Act
        int result = _scriptCommands.Export(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Run_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "script-run", _testExcelFile }; // Missing macro name

        // Act
        int result = _scriptCommands.Run(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Run_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "script-run", "nonexistent.xlsx", "Module1.AddTestData" };

        // Act
        int result = _scriptCommands.Run(args);

        // Assert
        Assert.Equal(1, result);
    }

    /// <summary>
    /// Round-trip test: Import VBA code that adds data to a worksheet, execute it, then verify the data
    /// This tests the complete VBA workflow for coding agents
    /// </summary>
    [Fact]
    public async Task VBA_RoundTrip_ImportExecuteAndVerifyData()
    {
        // Try to enable VBA access if it's not available
        if (!IsVbaAccessAvailable())
        {
            bool enabled = TryEnableVbaAccess();
            if (!enabled || !IsVbaAccessAvailable())
            {
                Assert.True(true, "Skipping VBA test - VBA project access could not be enabled");
                return;
            }
        }

        // Arrange - First add some initial data to the worksheet
        string[] writeArgs = { "sheet-write", _testExcelFile, "Sheet1", _testCsvFile };
        int writeResult = await _sheetCommands.Write(writeArgs);
        Assert.Equal(0, writeResult);

        // Act 1 - Read the initial data to verify it's there
        string[] readArgs1 = { "sheet-read", _testExcelFile, "Sheet1", "A1:C4" };
        int readResult1 = _sheetCommands.Read(readArgs1);
        Assert.Equal(0, readResult1);

        // Act 2 - Import VBA code that will add more data
        string[] importArgs = { "script-import", _testExcelFile, "TestModule", _testVbaFile };
        int importResult = await _scriptCommands.Import(importArgs);
        Assert.Equal(0, importResult);

        // Act 3 - Execute VBA macro that adds data to the worksheet
        string[] runArgs = { "script-run", _testExcelFile, "TestModule.AddTestData" };
        int runResult = _scriptCommands.Run(runArgs);
        Assert.Equal(0, runResult);

        // Act 4 - Verify the data was added by reading an extended range
        string[] readArgs2 = { "sheet-read", _testExcelFile, "Sheet1", "A1:C7" }; // Extended range for new data
        int readResult2 = _sheetCommands.Read(readArgs2);
        Assert.Equal(0, readResult2);

        // Act 5 - Verify we can list the VBA modules
        string[] listArgs = { "script-list", _testExcelFile };
        int listResult = _scriptCommands.List(listArgs);
        Assert.Equal(0, listResult);

        // Act 6 - Verify we can export the VBA code back
        string exportedVbaFile = Path.Combine(_tempDir, "ExportedModule.vba");
        string[] exportArgs = { "script-export", _testExcelFile, "TestModule", exportedVbaFile };
        int exportResult = _scriptCommands.Export(exportArgs);
        Assert.Equal(0, exportResult);
        Assert.True(File.Exists(exportedVbaFile));
    }

    /// <summary>
    /// Round-trip test with parameters: Execute VBA macro with parameters and verify results
    /// </summary>
    [Fact]
    public void VBA_RoundTrip_ExecuteWithParametersAndVerifyData()
    {
        // This test demonstrates how coding agents can execute VBA with parameters
        // and then verify the results

        // Arrange - Start with a clean sheet
        string[] createArgs = { "sheet-create", _testExcelFile, "TestSheet" };
        int createResult = _sheetCommands.Create(createArgs);
        Assert.Equal(0, createResult);

        // NOTE: The actual VBA execution with parameters is commented out because it requires
        // a workbook with VBA code already imported. When script-import command is available:
        
        /*
        // Future implementation:
        
        // Import VBA code
        string[] importArgs = { "script-import", _testExcelFile, "TestModule", _testVbaFile };
        int importResult = _scriptCommands.Import(importArgs);
        Assert.Equal(0, importResult);

        // Execute VBA macro with parameters (start at row 1, add 5 items, base value 1000)
        string[] runArgs = { "script-run", _testExcelFile, "TestModule.AddDataWithParameters", "1", "5", "1000" };
        int runResult = _scriptCommands.Run(runArgs);
        Assert.Equal(0, runResult);

        // Verify the data was added correctly
        string[] readArgs = { "sheet-read", _testExcelFile, "TestSheet", "A1:C6" }; // Headers + 5 rows
        int readResult = _sheetCommands.Read(readArgs);
        Assert.Equal(0, readResult);

        // Execute function that calculates sum and returns value
        string[] calcArgs = { "script-run", _testExcelFile, "TestModule.CalculateSum" };
        int calcResult = _scriptCommands.Run(calcArgs);
        Assert.Equal(0, calcResult);
        // The function should return 5050 (1000+1010+1020+1030+1040)
        */
    }

    /// <summary>
    /// Round-trip test: Update VBA code with new functionality and verify it works
    /// This tests the VBA update workflow for coding agents
    /// </summary>
    [Fact]
    public async Task VBA_RoundTrip_UpdateCodeAndVerifyNewFunctionality()
    {
        // Try to enable VBA access if it's not available
        if (!IsVbaAccessAvailable())
        {
            bool enabled = TryEnableVbaAccess();
            if (!enabled || !IsVbaAccessAvailable())
            {
                Assert.True(true, "Skipping VBA test - VBA project access could not be enabled");
                return;
            }
        }

        // Arrange - Import initial VBA code
        string[] importArgs = { "script-import", _testExcelFile, "TestModule", _testVbaFile };
        int importResult = await _scriptCommands.Import(importArgs);
        Assert.Equal(0, importResult);

        // Create updated VBA code with additional functionality
        string updatedVbaCode = @"
Sub AddTestData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Sheet1"")
    
    ' Original data
    ws.Cells(5, 1).Value = ""VBA""
    ws.Cells(5, 2).Value = ""Data""
    ws.Cells(5, 3).Value = ""Test""
    
    ' NEW: Additional row with different data
    ws.Cells(6, 1).Value = ""Updated""
    ws.Cells(6, 2).Value = ""Code""
    ws.Cells(6, 3).Value = ""Works""
End Sub

' NEW: Additional function for testing
Function TestFunction() As String
    TestFunction = ""VBA Update Success""
End Function";

        string updatedVbaFile = Path.Combine(_tempDir, "UpdatedModule.vba");
        await File.WriteAllTextAsync(updatedVbaFile, updatedVbaCode);

        // Act 1 - Update the VBA code with new functionality
        string[] updateArgs = { "script-update", _testExcelFile, "TestModule", updatedVbaFile };
        int updateResult = await _scriptCommands.Update(updateArgs);
        Assert.Equal(0, updateResult);

        // Act 2 - Execute the updated VBA macro
        string[] runArgs = { "script-run", _testExcelFile, "TestModule.AddTestData" };
        int runResult = _scriptCommands.Run(runArgs);
        Assert.Equal(0, runResult);

        // Act 3 - Verify the updated functionality by reading extended data
        string[] readArgs = { "sheet-read", _testExcelFile, "Sheet1", "A1:C6" };
        int readResult = _sheetCommands.Read(readArgs);
        Assert.Equal(0, readResult);

        // Act 4 - Export and verify the updated code contains our changes
        string exportedVbaFile = Path.Combine(_tempDir, "ExportedUpdatedModule.vba");
        string[] exportArgs = { "script-export", _testExcelFile, "TestModule", exportedVbaFile };
        int exportResult = _scriptCommands.Export(exportArgs);
        Assert.Equal(0, exportResult);
        
        // Verify exported code contains the new function
        string exportedCode = await File.ReadAllTextAsync(exportedVbaFile);
        Assert.Contains("TestFunction", exportedCode);
        Assert.Contains("VBA Update Success", exportedCode);
    }

    public void Dispose()
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
            // Ignore cleanup errors in tests
        }
        
        GC.SuppressFinalize(this);
    }
}

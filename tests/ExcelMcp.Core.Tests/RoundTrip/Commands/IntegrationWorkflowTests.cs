using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using System.IO;

namespace Sbroenne.ExcelMcp.Core.Tests.RoundTrip.Commands;

/// <summary>
/// Round trip tests for complete Core workflows combining multiple operations.
/// These tests require Excel installation and validate end-to-end Core data operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Workflows")]
public class IntegrationWorkflowTests : IDisposable
{
    private readonly IFileCommands _fileCommands;
    private readonly ISheetCommands _sheetCommands;
    private readonly ICellCommands _cellCommands;
    private readonly IParameterCommands _parameterCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public IntegrationWorkflowTests()
    {
        _fileCommands = new FileCommands();
        _sheetCommands = new SheetCommands();
        _cellCommands = new CellCommands();
        _parameterCommands = new ParameterCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_Integration_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");
        
        // Create test Excel file
        CreateTestExcelFile();
    }

    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmpty(_testExcelFile, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public void Workflow_CreateFile_AddSheet_WriteData_ReadData()
    {
        // 1. Validate file exists
        Assert.True(File.Exists(_testExcelFile), "Test Excel file should exist");

        // 2. Create new sheet
        var createSheetResult = _sheetCommands.Create(_testExcelFile, "DataSheet");
        Assert.True(createSheetResult.Success);

        // 3. Write data
        var csvPath = Path.Combine(_tempDir, "data.csv");
        File.WriteAllText(csvPath, "Name,Age\nAlice,30\nBob,25");
        var writeResult = _sheetCommands.Write(_testExcelFile, "DataSheet", csvPath);
        Assert.True(writeResult.Success);

        // 4. Read data back
        var readResult = _sheetCommands.Read(_testExcelFile, "DataSheet", "A1:B3");
        Assert.True(readResult.Success);
        Assert.NotEmpty(readResult.Data);
    }

    [Fact]
    public void Workflow_SetCellValue_CreateParameter_GetParameter()
    {
        // 1. Set cell value
        var setCellResult = _cellCommands.SetValue(_testExcelFile, "Sheet1", "A1", "TestValue");
        Assert.True(setCellResult.Success, $"Failed to set cell value: {setCellResult.ErrorMessage}");

        // 2. Create parameter (named range) pointing to cell - Use unique name
        string paramName = "TestParam_" + Guid.NewGuid().ToString("N")[..8];
        var createParamResult = _parameterCommands.Create(_testExcelFile, paramName, "Sheet1!A1");
        Assert.True(createParamResult.Success, $"Failed to create parameter: {createParamResult.ErrorMessage}");

        // 3. Get parameter value
        var getParamResult = _parameterCommands.Get(_testExcelFile, paramName);
        Assert.True(getParamResult.Success, $"Failed to get parameter: {getParamResult.ErrorMessage}");
        Assert.Equal("TestValue", getParamResult.Value?.ToString());
    }

    [Fact]
    public void Workflow_MultipleSheets_WithData_AndParameters()
    {
        // 1. Create multiple sheets
        var sheet1Result = _sheetCommands.Create(_testExcelFile, "Config");
        var sheet2Result = _sheetCommands.Create(_testExcelFile, "Data");
        Assert.True(sheet1Result.Success && sheet2Result.Success);

        // 2. Set configuration values
        _cellCommands.SetValue(_testExcelFile, "Config", "A1", "AppName");
        _cellCommands.SetValue(_testExcelFile, "Config", "B1", "MyApp");

        // 3. Create parameters - Use unique names
        string labelParam = "AppNameLabel_" + Guid.NewGuid().ToString("N")[..8];
        string valueParam = "AppNameValue_" + Guid.NewGuid().ToString("N")[..8];
        _parameterCommands.Create(_testExcelFile, labelParam, "Config!A1");
        _parameterCommands.Create(_testExcelFile, valueParam, "Config!B1");

        // 4. List parameters
        var listResult = _parameterCommands.List(_testExcelFile);
        Assert.True(listResult.Success);
        Assert.True(listResult.Parameters.Count >= 2);
    }

    [Fact]
    public void Workflow_CreateSheets_RenameSheet_DeleteSheet()
    {
        // 1. Create sheets
        _sheetCommands.Create(_testExcelFile, "Temp1");
        _sheetCommands.Create(_testExcelFile, "Temp2");

        // 2. Rename sheet
        var renameResult = _sheetCommands.Rename(_testExcelFile, "Temp1", "Renamed");
        Assert.True(renameResult.Success);

        // 3. Verify rename
        var listResult = _sheetCommands.List(_testExcelFile);
        Assert.Contains(listResult.Worksheets, w => w.Name == "Renamed");
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "Temp1");

        // 4. Delete sheet
        var deleteResult = _sheetCommands.Delete(_testExcelFile, "Temp2");
        Assert.True(deleteResult.Success);
    }

    [Fact]
    public void Workflow_SetFormula_GetFormula_ReadCalculatedValue()
    {
        // 1. Set values
        _cellCommands.SetValue(_testExcelFile, "Sheet1", "A1", "10");
        _cellCommands.SetValue(_testExcelFile, "Sheet1", "A2", "20");

        // 2. Set formula
        var setFormulaResult = _cellCommands.SetFormula(_testExcelFile, "Sheet1", "A3", "=SUM(A1:A2)");
        Assert.True(setFormulaResult.Success);

        // 3. Get formula
        var getFormulaResult = _cellCommands.GetFormula(_testExcelFile, "Sheet1", "A3");
        Assert.True(getFormulaResult.Success);
        Assert.Contains("SUM", getFormulaResult.Formula);

        // 4. Get calculated value
        var getValueResult = _cellCommands.GetValue(_testExcelFile, "Sheet1", "A3");
        Assert.True(getValueResult.Success);
        // Excel may return numeric value as number or string, so compare as string
        Assert.Equal("30", getValueResult.Value?.ToString());
    }

    [Fact]
    public void Workflow_CopySheet_ModifyOriginal_VerifyCopyUnchanged()
    {
        // 1. Set value in original sheet
        _cellCommands.SetValue(_testExcelFile, "Sheet1", "A1", "Original");

        // 2. Copy sheet
        var copyResult = _sheetCommands.Copy(_testExcelFile, "Sheet1", "Sheet1_Copy");
        Assert.True(copyResult.Success);

        // 3. Modify original
        _cellCommands.SetValue(_testExcelFile, "Sheet1", "A1", "Modified");

        // 4. Verify copy unchanged
        var copyValue = _cellCommands.GetValue(_testExcelFile, "Sheet1_Copy", "A1");
        Assert.Equal("Original", copyValue.Value);
    }

    [Fact]
    public void Workflow_AppendData_VerifyMultipleRows()
    {
        // 1. Initial write
        var csv1 = Path.Combine(_tempDir, "data1.csv");
        File.WriteAllText(csv1, "Name,Score\nAlice,90");
        _sheetCommands.Write(_testExcelFile, "Sheet1", csv1);

        // 2. Append more data
        var csv2 = Path.Combine(_tempDir, "data2.csv");
        File.WriteAllText(csv2, "Bob,85\nCharlie,95");
        var appendResult = _sheetCommands.Append(_testExcelFile, "Sheet1", csv2);
        Assert.True(appendResult.Success);

        // 3. Read all data
        var readResult = _sheetCommands.Read(_testExcelFile, "Sheet1", "A1:B4");
        Assert.True(readResult.Success);
        Assert.Equal(4, readResult.Data.Count); // Header + 3 data rows
    }

    [Fact]
    public void Workflow_ClearRange_VerifyEmptyCells()
    {
        // 1. Write data
        _cellCommands.SetValue(_testExcelFile, "Sheet1", "A1", "Data1");
        _cellCommands.SetValue(_testExcelFile, "Sheet1", "A2", "Data2");

        // 2. Clear range
        var clearResult = _sheetCommands.Clear(_testExcelFile, "Sheet1", "A1:A2");
        Assert.True(clearResult.Success);

        // 3. Verify cleared
        var value1 = _cellCommands.GetValue(_testExcelFile, "Sheet1", "A1");
        var value2 = _cellCommands.GetValue(_testExcelFile, "Sheet1", "A2");
        Assert.True(value1.Value == null || string.IsNullOrEmpty(value1.Value.ToString()));
        Assert.True(value2.Value == null || string.IsNullOrEmpty(value2.Value.ToString()));
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

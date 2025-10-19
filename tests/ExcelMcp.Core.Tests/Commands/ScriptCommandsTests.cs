using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using System.IO;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Script (VBA) Core operations.
/// These tests require Excel installation and VBA trust enabled.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public class ScriptCommandsTests : IDisposable
{
    private readonly IScriptCommands _scriptCommands;
    private readonly IFileCommands _fileCommands;
    private readonly ISetupCommands _setupCommands;
    private readonly string _testExcelFile;
    private readonly string _testVbaFile;
    private readonly string _tempDir;
    private bool _disposed;

    public ScriptCommandsTests()
    {
        _scriptCommands = new ScriptCommands();
        _fileCommands = new FileCommands();
        _setupCommands = new SetupCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_VBA_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsm");
        _testVbaFile = Path.Combine(_tempDir, "TestModule.vba");
        
        // Create test files
        CreateTestExcelFile();
        CreateTestVbaFile();
        
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

    private void CreateTestVbaFile()
    {
        string vbaCode = @"Option Explicit

Public Function TestFunction() As String
    TestFunction = ""Hello from VBA""
End Function

Public Sub TestSubroutine()
    MsgBox ""Test VBA""
End Sub";
    
        File.WriteAllText(_testVbaFile, vbaCode);
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
    public void List_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _scriptCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Scripts);
        Assert.Empty(result.Scripts); // New file has no VBA modules
    }

    [Fact]
    public async Task Import_WithValidVbaCode_ReturnsSuccessResult()
    {
        // Act
        var result = await _scriptCommands.Import(_testExcelFile, "TestModule", _testVbaFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    [Fact]
    public async Task List_AfterImport_ShowsNewModule()
    {
        // Arrange
        await _scriptCommands.Import(_testExcelFile, "TestModule", _testVbaFile);

        // Act
        var result = _scriptCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Scripts);
        Assert.Single(result.Scripts);
        Assert.Equal("TestModule", result.Scripts[0].Name);
    }

    [Fact]
    public async Task Export_WithExistingModule_CreatesFile()
    {
        // Arrange
        await _scriptCommands.Import(_testExcelFile, "TestModule", _testVbaFile);
        var exportPath = Path.Combine(_tempDir, "exported.vba");

        // Act
        var result = await _scriptCommands.Export(_testExcelFile, "TestModule", exportPath);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public async Task Update_WithValidVbaCode_ReturnsSuccessResult()
    {
        // Arrange
        await _scriptCommands.Import(_testExcelFile, "TestModule", _testVbaFile);
        var updateFile = Path.Combine(_tempDir, "updated.vba");
        File.WriteAllText(updateFile, "Public Function Updated() As String\n    Updated = \"Updated\"\nEnd Function");

        // Act
        var result = await _scriptCommands.Update(_testExcelFile, "TestModule", updateFile);

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task Delete_WithExistingModule_ReturnsSuccessResult()
    {
        // Arrange
        await _scriptCommands.Import(_testExcelFile, "TestModule", _testVbaFile);

        // Act
        var result = _scriptCommands.Delete(_testExcelFile, "TestModule");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task Import_ThenDelete_ThenList_ShowsEmpty()
    {
        // Arrange
        await _scriptCommands.Import(_testExcelFile, "TestModule", _testVbaFile);
        _scriptCommands.Delete(_testExcelFile, "TestModule");

        // Act
        var result = _scriptCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.Empty(result.Scripts);
    }

    [Fact]
    public async Task Export_WithNonExistentModule_ReturnsErrorResult()
    {
        // Arrange
        var exportPath = Path.Combine(_tempDir, "nonexistent.vba");

        // Act
        var result = await _scriptCommands.Export(_testExcelFile, "NonExistentModule", exportPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
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

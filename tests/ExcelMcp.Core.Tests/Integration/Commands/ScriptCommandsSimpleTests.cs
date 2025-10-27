using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Simple integration tests for ScriptCommands using batch pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "VBA")]
[Trait("RequiresExcel", "true")]
public class ScriptCommandsSimpleTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFile;
    private readonly ScriptCommands _commands;
    private readonly FileCommands _fileCommands;

    public ScriptCommandsSimpleTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_ScriptSimple_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _testFile = Path.Combine(_testDir, "test.xlsm");
        _commands = new ScriptCommands();
        _fileCommands = new FileCommands();

        // Create test workbook (xlsm extension creates macro-enabled file)
        var result = _fileCommands.CreateEmptyAsync(_testFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test file: {result.ErrorMessage}");
        }
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, recursive: true);
            }
        }
        catch { /* Cleanup failure is non-critical */ }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task List_EmptyMacroFile_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Scripts);
        // New macro file may have default modules like "ThisWorkbook"
    }

    [Fact]
    public async Task Import_BasicModule_Success()
    {
        // Arrange
        const string moduleName = "TestModule";
        var vbaCodeFile = Path.Combine(_testDir, $"{moduleName}.bas");
        File.WriteAllText(vbaCodeFile, @"Attribute VB_Name = ""TestModule""
Sub TestSub()
    ' Simple test sub
End Sub
");

        // Act - Import module
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var importResult = await _commands.ImportAsync(batch, moduleName, vbaCodeFile);
            Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");
            await batch.SaveAsync();
        }

        // Act - List modules (new batch)
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var listResult = await _commands.ListAsync(batch);

            // Assert
            Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
            Assert.Contains(listResult.Scripts!, m => m.Name == moduleName);
        }
    }
}

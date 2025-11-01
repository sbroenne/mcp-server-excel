using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.VbaTrust;

/// <summary>
/// Tests for VBA operations when trust is enabled (CI environment has trust enabled)
/// </summary>
public partial class VbaTrustDetectionTests
{
    [Fact]
    public async Task ScriptCommands_List_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_List_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _scriptCommands.ListAsync(batch);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"List should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Scripts);
    }

    [Fact]
    public async Task ScriptCommands_Import_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Import_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        File.WriteAllText(vbaFile, "Sub TestImport()\nEnd Sub");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"Import should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
    }

    [Fact]
    public async Task ScriptCommands_Export_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Export_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // First import a module so we have something to export
        string vbaFile = Path.Join(_tempDir, $"ImportModule_{Guid.NewGuid():N}.vba");
        File.WriteAllText(vbaFile, "Sub TestCode()\nEnd Sub");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success, "Import should succeed before export test");

        string exportFile = Path.Join(_tempDir, $"ExportedModule_{Guid.NewGuid():N}.vba");

        // Act - Export the module we just imported
        var result = await _scriptCommands.ExportAsync(batch, "TestModule", exportFile);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"Export should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.True(File.Exists(exportFile), "Exported file should exist");
    }

    [Fact]
    public async Task ScriptCommands_Run_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Run_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a test macro first
        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        string vbaCode = @"Sub TestProcedure()
    ' Simple test procedure
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success);

        // Act - Run the macro
        var runResult = await _scriptCommands.RunAsync(batch, "TestModule.TestProcedure");

        // Assert - Should succeed when VBA trust is enabled
        Assert.True(runResult.Success, $"Run should succeed with VBA trust enabled. Error: {runResult.ErrorMessage}");
    }

    [Fact]
    public async Task ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a test module first
        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        File.WriteAllText(vbaFile, "Sub Test()\nEnd Sub");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success);

        // Act - Delete the module
        var deleteResult = await _scriptCommands.DeleteAsync(batch, "TestModule");

        // Assert - Should succeed when VBA trust is enabled
        Assert.True(deleteResult.Success, $"Delete should succeed with VBA trust enabled. Error: {deleteResult.ErrorMessage}");
    }
}

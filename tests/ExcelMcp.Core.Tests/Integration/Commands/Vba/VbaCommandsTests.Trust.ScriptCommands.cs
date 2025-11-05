using System.IO;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Vba;

/// <summary>
/// Tests for VBA operations when trust is enabled (CI environment has trust enabled)
/// </summary>
public partial class VbaCommandsTests
{
    [Fact]
    public async Task ScriptCommands_List_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaCommandsTests), nameof(ScriptCommands_List_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

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
            nameof(VbaCommandsTests), nameof(ScriptCommands_Import_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        System.IO.File.WriteAllText(vbaFile, "Sub TestImport()\nEnd Sub");

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
            nameof(VbaCommandsTests), nameof(ScriptCommands_Export_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // First import a module so we have something to export
        string vbaFile = Path.Join(_tempDir, $"ImportModule_{Guid.NewGuid():N}.vba");
        System.IO.File.WriteAllText(vbaFile, "Sub TestCode()\nEnd Sub");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success, "Import should succeed before export test");

        string exportFile = Path.Join(_tempDir, $"ExportedModule_{Guid.NewGuid():N}.vba");

        // Act - Export the module we just imported
        var result = await _scriptCommands.ExportAsync(batch, "TestModule", exportFile);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"Export should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.True(System.IO.File.Exists(exportFile), "Exported file should exist");
    }

    [Fact]
    public async Task ScriptCommands_Run_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Run_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a test macro first
        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        string vbaCode = @"Sub TestProcedure()
    ' Simple test procedure
End Sub";
        System.IO.File.WriteAllText(vbaFile, vbaCode);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success);

        // Act - Run the macro
        var runResult = await _scriptCommands.RunAsync(batch, "TestModule.TestProcedure", null);

        // Assert - Should succeed when VBA trust is enabled
        Assert.True(runResult.Success, $"Run should succeed with VBA trust enabled. Error: {runResult.ErrorMessage}");
    }

    [Fact]
    public async Task ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a module first
        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        System.IO.File.WriteAllText(vbaFile, "Sub TestCode()\nEnd Sub");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success);

        // Act - Delete the module
        var result = await _scriptCommands.DeleteAsync(batch, "TestModule");

        // Assert - Should succeed when VBA trust is enabled
        Assert.True(result.Success, $"Delete should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        
        // Verify module is gone
        var listResult = await _scriptCommands.ListAsync(batch);
        Assert.DoesNotContain(listResult.Scripts, s => s.Name == "TestModule");
    }

    [Fact]
    public async Task ScriptCommands_View_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaCommandsTests), nameof(ScriptCommands_View_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a module with known code
        string vbaFile = Path.Join(_tempDir, $"ViewTestModule_{Guid.NewGuid():N}.vba");
        string expectedCode = "Sub ViewTest()\n    MsgBox \"Hello\"\nEnd Sub";
        System.IO.File.WriteAllText(vbaFile, expectedCode);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "ViewTestModule", vbaFile);
        Assert.True(importResult.Success, "Import should succeed before view test");

        // Act - View the module code
        var result = await _scriptCommands.ViewAsync(batch, "ViewTestModule");

        // Assert - Should succeed and return the code
        Assert.True(result.Success, $"View should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Code);
        Assert.Contains("ViewTest", result.Code);
        Assert.Contains("MsgBox", result.Code);
    }

    [Fact]
    public async Task ScriptCommands_Update_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Update_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import initial module
        string vbaFile = Path.Join(_tempDir, $"UpdateTestModule_{Guid.NewGuid():N}.vba");
        string initialCode = "Sub OriginalCode()\nEnd Sub";
        System.IO.File.WriteAllText(vbaFile, initialCode);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var importResult = await _scriptCommands.ImportAsync(batch, "UpdateTestModule", vbaFile);
        Assert.True(importResult.Success, "Import should succeed before update test");

        // Prepare updated code
        string updatedVbaFile = Path.Join(_tempDir, $"UpdatedModule_{Guid.NewGuid():N}.vba");
        string updatedCode = "Sub UpdatedCode()\n    MsgBox \"Updated\"\nEnd Sub";
        System.IO.File.WriteAllText(updatedVbaFile, updatedCode);

        // Act - Update the module with new code
        var result = await _scriptCommands.UpdateAsync(batch, "UpdateTestModule", updatedVbaFile);

        // Assert - Should succeed
        Assert.True(result.Success, $"Update should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        
        // Verify the code was updated
        var viewResult = await _scriptCommands.ViewAsync(batch, "UpdateTestModule");
        Assert.True(viewResult.Success);
        Assert.Contains("UpdatedCode", viewResult.Code);
        Assert.Contains("Updated", viewResult.Code);
        Assert.DoesNotContain("OriginalCode", viewResult.Code);
    }
}

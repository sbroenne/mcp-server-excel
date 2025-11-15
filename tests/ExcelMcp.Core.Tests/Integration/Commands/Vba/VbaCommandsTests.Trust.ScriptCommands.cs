using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Vba;

/// <summary>
/// Tests for VBA operations when trust is enabled (CI environment has trust enabled)
/// </summary>
public partial class VbaCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public async Task ScriptCommands_List_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(VbaCommandsTests), nameof(ScriptCommands_List_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _scriptCommands.List(batch);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"List should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Scripts);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ScriptCommands_Import_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Import_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        System.IO.File.WriteAllText(vbaFile, "Sub TestImport()\nEnd Sub");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _scriptCommands.Import(batch, "TestModule", vbaFile);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"Import should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ScriptCommands_Export_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Export_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // First import a module so we have something to export
        string vbaFile = Path.Join(_tempDir, $"ImportModule_{Guid.NewGuid():N}.vba");
        System.IO.File.WriteAllText(vbaFile, "Sub TestCode()\nEnd Sub");

        using var batch = ExcelSession.BeginBatch(testFile);
        var importResult = _scriptCommands.Import(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success, "Import should succeed before export test");

        // Act - View (export) the module we just imported
        var result = _scriptCommands.View(batch, "TestModule");

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"View should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Code);
        Assert.NotEmpty(result.Code);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ScriptCommands_Run_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Run_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a test macro first
        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        string vbaCode = @"Sub TestProcedure()
    ' Simple test procedure
End Sub";
        System.IO.File.WriteAllText(vbaFile, vbaCode);

        using var batch = ExcelSession.BeginBatch(testFile);
        var importResult = _scriptCommands.Import(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success);

        // Act - Run the macro
        var runResult = _scriptCommands.Run(batch, "TestModule.TestProcedure", null);

        // Assert - Should succeed when VBA trust is enabled
        Assert.True(runResult.Success, $"Run should succeed with VBA trust enabled. Error: {runResult.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a module first
        string vbaFile = Path.Join(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        System.IO.File.WriteAllText(vbaFile, "Sub TestCode()\nEnd Sub");

        using var batch = ExcelSession.BeginBatch(testFile);
        var importResult = _scriptCommands.Import(batch, "TestModule", vbaFile);
        Assert.True(importResult.Success);

        // Act - Delete the module
        var result = _scriptCommands.Delete(batch, "TestModule");

        // Assert - Should succeed when VBA trust is enabled
        Assert.True(result.Success, $"Delete should succeed with VBA trust enabled. Error: {result.ErrorMessage}");

        // Verify module is gone
        var listResult = _scriptCommands.List(batch);
        Assert.DoesNotContain(listResult.Scripts, s => s.Name == "TestModule");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ScriptCommands_View_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(VbaCommandsTests), nameof(ScriptCommands_View_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import a module with known code
        string vbaFile = Path.Join(_tempDir, $"ViewTestModule_{Guid.NewGuid():N}.vba");
        string expectedCode = "Sub ViewTest()\n    MsgBox \"Hello\"\nEnd Sub";
        System.IO.File.WriteAllText(vbaFile, expectedCode);

        using var batch = ExcelSession.BeginBatch(testFile);
        var importResult = _scriptCommands.Import(batch, "ViewTestModule", vbaFile);
        Assert.True(importResult.Success, "Import should succeed before view test");

        // Act - View the module code
        var result = _scriptCommands.View(batch, "ViewTestModule");

        // Assert - Should succeed and return the code
        Assert.True(result.Success, $"View should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Code);
        Assert.Contains("ViewTest", result.Code);
        Assert.Contains("MsgBox", result.Code);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ScriptCommands_Update_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(VbaCommandsTests), nameof(ScriptCommands_Update_WithTrustEnabled_WorksCorrectly), _tempDir, ".xlsm");

        // Import initial module
        string vbaFile = Path.Join(_tempDir, $"UpdateTestModule_{Guid.NewGuid():N}.vba");
        string initialCode = "Sub OriginalCode()\nEnd Sub";
        System.IO.File.WriteAllText(vbaFile, initialCode);

        using var batch = ExcelSession.BeginBatch(testFile);
        var importResult = _scriptCommands.Import(batch, "UpdateTestModule", vbaFile);
        Assert.True(importResult.Success, "Import should succeed before update test");

        // Prepare updated code
        string updatedVbaFile = Path.Join(_tempDir, $"UpdatedModule_{Guid.NewGuid():N}.vba");
        string updatedCode = "Sub UpdatedCode()\n    MsgBox \"Updated\"\nEnd Sub";
        System.IO.File.WriteAllText(updatedVbaFile, updatedCode);

        // Act - Update the module with new code
        var result = _scriptCommands.Update(batch, "UpdateTestModule", updatedVbaFile);

        // Assert - Should succeed
        Assert.True(result.Success, $"Update should succeed with VBA trust enabled. Error: {result.ErrorMessage}");

        // Verify the code was updated
        var viewResult = _scriptCommands.View(batch, "UpdateTestModule");
        Assert.True(viewResult.Success);
        Assert.Contains("UpdatedCode", viewResult.Code);
        Assert.Contains("Updated", viewResult.Code);
        Assert.DoesNotContain("OriginalCode", viewResult.Code);
    }
}

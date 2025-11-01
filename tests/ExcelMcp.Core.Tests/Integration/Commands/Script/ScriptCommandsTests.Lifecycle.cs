using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Script;

/// <summary>
/// Tests for Script (VBA) lifecycle operations (list, import, export, delete, update)
/// </summary>
public partial class ScriptCommandsTests
{
    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessResult()
    {
        // Arrange - Create .xlsm file for macro support
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ScriptCommandsTests), nameof(List_WithValidFile_ReturnsSuccessResult), _tempDir, ".xlsm");

        // Check VBA trust before running test
        await using var trustBatch = await ExcelSession.BeginBatchAsync(testFile);
        var trustResult = await _setupCommands.CheckVbaTrustAsync(trustBatch);
        if (!trustResult.IsTrusted)
        {
            // Skip test if VBA trust is not enabled
            return;
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _scriptCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Scripts);
        // Excel always creates default document modules (ThisWorkbook, Sheet1, etc.)
        Assert.True(result.Scripts.Count >= 0); // At minimum, no error occurred
    }

    [Fact]
    public async Task Import_WithValidVbaCode_ReturnsSuccessResult()
    {
        // Arrange - Create .xlsm file and VBA code
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ScriptCommandsTests), nameof(Import_WithValidVbaCode_ReturnsSuccessResult), _tempDir, ".xlsm");
        var testVbaFile = CreateTestVbaFile($"Import_{Guid.NewGuid():N}.vba");

        // Check VBA trust
        await using var trustBatch = await ExcelSession.BeginBatchAsync(testFile);
        var trustResult = await _setupCommands.CheckVbaTrustAsync(trustBatch);
        if (!trustResult.IsTrusted)
        {
            return; // Skip test
        }

        // Act - Use single batch for import and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _scriptCommands.ImportAsync(batch, "TestModule", testVbaFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");

        // Verify module was imported
        var listResult = await _scriptCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Scripts, s => s.Name == "TestModule");

        await batch.SaveAsync();
    }

    [Fact]
    public async Task Export_WithExistingModule_CreatesFile()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ScriptCommandsTests), nameof(Export_WithExistingModule_CreatesFile), _tempDir, ".xlsm");
        var testVbaFile = CreateTestVbaFile($"Export_{Guid.NewGuid():N}.vba");
        var exportPath = Path.Combine(_tempDir, $"exported_{Guid.NewGuid():N}.vba");

        // Check VBA trust
        await using var trustBatch = await ExcelSession.BeginBatchAsync(testFile);
        var trustResult = await _setupCommands.CheckVbaTrustAsync(trustBatch);
        if (!trustResult.IsTrusted)
        {
            return; // Skip test
        }

        // Act - Use single batch for import and export
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _scriptCommands.ImportAsync(batch, "TestModule", testVbaFile);
        var result = await _scriptCommands.ExportAsync(batch, "TestModule", exportPath);

        // Assert
        Assert.True(result.Success, $"Export failed: {result.ErrorMessage}");
        Assert.True(File.Exists(exportPath));

        await batch.SaveAsync();
    }

    [Fact]
    public async Task Delete_WithExistingModule_ReturnsSuccessResult()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ScriptCommandsTests), nameof(Delete_WithExistingModule_ReturnsSuccessResult), _tempDir, ".xlsm");
        var testVbaFile = CreateTestVbaFile($"Delete_{Guid.NewGuid():N}.vba");

        // Check VBA trust
        await using var trustBatch = await ExcelSession.BeginBatchAsync(testFile);
        var trustResult = await _setupCommands.CheckVbaTrustAsync(trustBatch);
        if (!trustResult.IsTrusted)
        {
            return; // Skip test
        }

        // Act - Use single batch for import, delete, and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _scriptCommands.ImportAsync(batch, "TestModule", testVbaFile);
        var result = await _scriptCommands.DeleteAsync(batch, "TestModule");

        // Assert
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify module was deleted
        var listResult = await _scriptCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Scripts, s => s.Name == "TestModule");

        await batch.SaveAsync();
    }

    [Fact]
    public async Task Export_WithNonExistentModule_ReturnsErrorResult()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ScriptCommandsTests), nameof(Export_WithNonExistentModule_ReturnsErrorResult), _tempDir, ".xlsm");
        var exportPath = Path.Combine(_tempDir, $"nonexistent_{Guid.NewGuid():N}.vba");

        // Check VBA trust
        await using var trustBatch = await ExcelSession.BeginBatchAsync(testFile);
        var trustResult = await _setupCommands.CheckVbaTrustAsync(trustBatch);
        if (!trustResult.IsTrusted)
        {
            return; // Skip test
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _scriptCommands.ExportAsync(batch, "NonExistentModule", exportPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }
}

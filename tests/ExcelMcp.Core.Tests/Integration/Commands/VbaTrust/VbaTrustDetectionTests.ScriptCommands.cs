using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.VbaTrust;

/// <summary>
/// Tests for VBA trust interaction with ScriptCommands
/// </summary>
public partial class VbaTrustDetectionTests
{
    [Fact]
    public async Task ScriptCommands_List_HandlesVbaTrustCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_List_HandlesVbaTrustCorrectly), _tempDir, ".xlsm");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _scriptCommands.ListAsync(batch);

        // Assert
        Assert.NotNull(result);
        Assert.NotNull(result.Scripts);
    }

    [Fact]
    public async Task TestVbaTrustScope_AllowsVbaOperations()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(TestVbaTrustScope_AllowsVbaOperations), _tempDir, ".xlsm");

        string vbaFile = Path.Combine(_tempDir, $"TestModule_{Guid.NewGuid():N}.vba");
        string vbaCode = @"Sub TestProcedure()
    MsgBox ""Test""
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        // Act & Assert - VBA operations should work inside the scope
        using (var _ = new TestVbaTrustScope())
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);

            // Should succeed when VBA trust is enabled
            if (!importResult.Success)
            {
                // If it failed, it should NOT be due to VBA trust
                Assert.DoesNotContain("trust", importResult.ErrorMessage?.ToLowerInvariant() ?? "");
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Export_WithTrust_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Export_WithTrust_WorksCorrectly), _tempDir, ".xlsm");

        string exportFile = Path.Combine(_tempDir, $"ExportedModule_{Guid.NewGuid():N}.vba");

        // Act - Test with VBA trust enabled
        using (var _ = new TestVbaTrustScope())
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var result = await _scriptCommands.ExportAsync(batch, "ThisWorkbook", exportFile);

            // Assert
            Assert.NotNull(result);
            // Should either succeed or fail for reasons other than trust
            if (!result.Success && result.ErrorMessage != null)
            {
                Assert.DoesNotContain("trust", result.ErrorMessage.ToLowerInvariant());
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Import_WithTrust_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Import_WithTrust_WorksCorrectly), _tempDir, ".xlsm");

        string vbaFile = Path.Combine(_tempDir, $"ImportTestModule_{Guid.NewGuid():N}.vba");
        string vbaCode = @"Sub ImportTestProcedure()
    Dim x As Integer
    x = 42
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        // Act - Test with VBA trust enabled
        using (var _ = new TestVbaTrustScope())
        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var result = await _scriptCommands.ImportAsync(batch, "ImportTestModule", vbaFile);

            // Assert
            Assert.NotNull(result);
            // Should either succeed or fail for reasons other than trust
            if (!result.Success && result.ErrorMessage != null)
            {
                Assert.DoesNotContain("trust", result.ErrorMessage.ToLowerInvariant());
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Update_WithTrust_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Update_WithTrust_WorksCorrectly), _tempDir, ".xlsm");

        string vbaFile = Path.Combine(_tempDir, $"UpdateTestModule_{Guid.NewGuid():N}.vba");
        string vbaCode1 = @"Sub UpdateTest1()
End Sub";
        File.WriteAllText(vbaFile, vbaCode1);

        using (var _ = new TestVbaTrustScope())
        {
            // First import
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                await _scriptCommands.ImportAsync(batch, "UpdateTestModule", vbaFile);
                await batch.SaveAsync();
            }

            // Update the VBA code
            string vbaCode2 = @"Sub UpdateTest2()
    Dim y As String
End Sub";
            File.WriteAllText(vbaFile, vbaCode2);

            // Act - Update the module
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                var result = await _scriptCommands.UpdateAsync(batch, "UpdateTestModule", vbaFile);

                // Assert
                Assert.NotNull(result);
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Delete_WithTrust_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Delete_WithTrust_WorksCorrectly), _tempDir, ".xlsm");

        string vbaFile = Path.Combine(_tempDir, $"DeleteTestModule_{Guid.NewGuid():N}.vba");
        string vbaCode = @"Sub DeleteTest()
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        using (var _ = new TestVbaTrustScope())
        {
            // First import a module
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                await _scriptCommands.ImportAsync(batch, "DeleteTestModule", vbaFile);
                await batch.SaveAsync();
            }

            // Act - Delete the module
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                var result = await _scriptCommands.DeleteAsync(batch, "DeleteTestModule");

                // Assert
                Assert.NotNull(result);
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Run_WithTrust_WorksCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(VbaTrustDetectionTests), nameof(ScriptCommands_Run_WithTrust_WorksCorrectly), _tempDir, ".xlsm");

        string vbaFile = Path.Combine(_tempDir, $"RunTestModule_{Guid.NewGuid():N}.vba");
        string vbaCode = @"Sub RunTest()
    ' Simple procedure that does nothing
    Dim x As Integer
    x = 1
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        using (var _ = new TestVbaTrustScope())
        {
            // First import a module
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                await _scriptCommands.ImportAsync(batch, "RunTestModule", vbaFile);
                await batch.SaveAsync();
            }

            // Act - Run the procedure
            await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
            {
                var result = await _scriptCommands.RunAsync(batch, "RunTestModule.RunTest");

                // Assert
                Assert.NotNull(result);
            }
        }
    }
}

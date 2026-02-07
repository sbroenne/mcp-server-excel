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
    public void ScriptCommands_List_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _scriptCommands.List(batch);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"List should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Scripts);
    }
    [Fact]
    public void ScriptCommands_Import_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string vbaCode = "Sub TestImport()\nEnd Sub";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        _scriptCommands.Import(batch, "TestModule", vbaCode);

        // Assert - verify module exists via list
        var listResult = _scriptCommands.List(batch);
        Assert.Contains(listResult.Scripts, s => s.Name == "TestModule");
    }
    [Fact]
    public void ScriptCommands_Export_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // First import a module so we have something to export
        string vbaCode = "Sub TestCode()\nEnd Sub";

        using var batch = ExcelSession.BeginBatch(testFile);
        _scriptCommands.Import(batch, "TestModule", vbaCode);

        // Act - View (export) the module we just imported
        var result = _scriptCommands.View(batch, "TestModule");

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"View should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Code);
        Assert.NotEmpty(result.Code);
    }
    [Fact]
    public void ScriptCommands_Run_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Import a test macro first
        string vbaCode = @"Sub TestProcedure()
    ' Simple test procedure
End Sub";

        using var batch = ExcelSession.BeginBatch(testFile);
        _scriptCommands.Import(batch, "TestModule", vbaCode);

        // Act - Run the macro
        _scriptCommands.Run(batch, "TestModule.TestProcedure", null);

        // Assert - No exception thrown; to be thorough, ensure module still exists
        var listResult = _scriptCommands.List(batch);
        Assert.Contains(listResult.Scripts, s => s.Name == "TestModule");
    }
    [Fact]
    public void ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Import a module first
        string vbaCode = "Sub TestCode()\nEnd Sub";

        using var batch = ExcelSession.BeginBatch(testFile);
        _scriptCommands.Import(batch, "TestModule", vbaCode);

        // Act - Delete the module
        _scriptCommands.Delete(batch, "TestModule");

        // Verify module is gone
        var listResult = _scriptCommands.List(batch);
        Assert.DoesNotContain(listResult.Scripts, s => s.Name == "TestModule");
    }
    [Fact]
    public void ScriptCommands_View_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Import a module with known code
        string expectedCode = "Sub ViewTest()\n    MsgBox \"Hello\"\nEnd Sub";

        using var batch = ExcelSession.BeginBatch(testFile);
        _scriptCommands.Import(batch, "ViewTestModule", expectedCode);

        // Act - View the module code
        var result = _scriptCommands.View(batch, "ViewTestModule");

        // Assert - Should succeed and return the code
        Assert.True(result.Success, $"View should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Code);
        Assert.Contains("ViewTest", result.Code);
        Assert.Contains("MsgBox", result.Code);
    }
    [Fact]
    public void ScriptCommands_Update_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Import initial module
        string initialCode = "Sub OriginalCode()\nEnd Sub";

        using var batch = ExcelSession.BeginBatch(testFile);
        _scriptCommands.Import(batch, "UpdateTestModule", initialCode);

        // Prepare updated code
        string updatedCode = "Sub UpdatedCode()\n    MsgBox \"Updated\"\nEnd Sub";

        // Act - Update the module with new code
        _scriptCommands.Update(batch, "UpdateTestModule", updatedCode);

        // Verify the code was updated
        var viewResult = _scriptCommands.View(batch, "UpdateTestModule");
        Assert.True(viewResult.Success);
        Assert.Contains("UpdatedCode", viewResult.Code);
        Assert.Contains("Updated", viewResult.Code);
        Assert.DoesNotContain("OriginalCode", viewResult.Code);
    }
}





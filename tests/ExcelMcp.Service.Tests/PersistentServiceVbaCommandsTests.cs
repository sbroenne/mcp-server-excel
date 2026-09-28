using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for VBA operations when trust is enabled (CI environment has trust enabled)
/// </summary>
internal interface IPersistentVbaCommands : IVbaCommands;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class PersistentServiceVbaCommandsTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceVbaFixture>
{
    private readonly IPersistentVbaCommands _scriptCommands;

    public PersistentServiceVbaCommandsTests(
        PersistentServiceVbaFixture fixture) :
        base(fixture)
    {
        _scriptCommands = fixture.CreateCommands<IPersistentVbaCommands>();
    }
    [Fact]
    public void ScriptCommands_List_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // Act
        var batch = _fixture.BatchToken;
        var result = _scriptCommands.List(batch);

        // Assert - Should succeed when VBA trust is enabled (as in CI environment)
        Assert.True(result.Success, $"List should succeed with VBA trust enabled. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Scripts);
    }
    [Fact]
    public void ScriptCommands_Import_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        string vbaCode = "Sub TestImport()\nEnd Sub";

        // Act
        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

        // Assert - verify module exists via list
        var listResult = _scriptCommands.List(batch);
        Assert.Contains(listResult.Scripts, s => s.Name == "TestModule");
    }
    [Fact]
    public void ScriptCommands_Export_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // First import a module so we have something to export
        string vbaCode = "Sub TestCode()\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

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
        // Import a test macro first
        string vbaCode = @"Sub TestProcedure()
    ' Simple test procedure
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

        // Act - Run the macro
        _scriptCommands.Run(batch, "TestModule.TestProcedure", null);

        // Assert - No exception thrown; to be thorough, ensure module still exists
        var listResult = _scriptCommands.List(batch);
        Assert.Contains(listResult.Scripts, s => s.Name == "TestModule");
    }

    [Fact]
    public async Task ScriptCommands_Run_OnReopenedMacroWorkbook_AfterList_WorksCorrectly()
    {
        // Arrange
        var rangeCommands = _commands;
        string vbaCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""reopened-run-ok""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "ReopenTestModule", vbaCode);
        await _fixture.SaveAndReopenAsync();

        // Act - reopen the existing .xlsm, prove VBA project access still works, then run
        var reopenedBatch = _fixture.BatchToken;
        var listResult = _scriptCommands.List(reopenedBatch);
        Assert.True(listResult.Success, $"List should succeed after reopen. Error: {listResult.ErrorMessage}");
        Assert.Contains(listResult.Scripts, s => s.Name == "ReopenTestModule");

        _scriptCommands.Run(reopenedBatch, "ReopenTestModule.WriteMarker", null);

        // Assert - macro execution against the reopened workbook should have real side effects
        var cellResult = rangeCommands.GetValues(reopenedBatch, "Sheet1", "A1");
        Assert.True(cellResult.Success, $"GetValues should succeed after reopened run. Error: {cellResult.ErrorMessage}");
        Assert.NotNull(cellResult.Values);
        Assert.Single(cellResult.Values);
        Assert.Equal("reopened-run-ok", cellResult.Values[0][0]?.ToString());
    }

    [Fact]
    public async Task ScriptCommands_Update_OnReopenedMacroWorkbook_ThenRun_UsesUpdatedCode()
    {
        // Arrange
        var rangeCommands = _commands;
        const string moduleName = "ReopenUpdateModule";
        string initialCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""original-run""
End Sub";
        string updatedCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""updated-run-ok""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, initialCode);
        await _fixture.SaveAndReopenAsync();

        // Act
        var reopenedBatch = _fixture.BatchToken;
        _scriptCommands.Update(reopenedBatch, moduleName, updatedCode);
        var viewResult = _scriptCommands.View(reopenedBatch, moduleName);
        _scriptCommands.Run(reopenedBatch, $"{moduleName}.WriteMarker", null);

        // Assert
        Assert.True(viewResult.Success, $"View should succeed after reopened update. Error: {viewResult.ErrorMessage}");
        Assert.Contains("updated-run-ok", viewResult.Code);
        Assert.DoesNotContain("original-run", viewResult.Code);

        var cellResult = rangeCommands.GetValues(reopenedBatch, "Sheet1", "A1");
        Assert.True(cellResult.Success, $"GetValues should succeed after reopened update run. Error: {cellResult.ErrorMessage}");
        Assert.NotNull(cellResult.Values);
        Assert.Single(cellResult.Values);
        Assert.Equal("updated-run-ok", cellResult.Values[0][0]?.ToString());
    }

    [Fact]
    public async Task ScriptCommands_DeleteThenImport_OnReopenedMacroWorkbook_ThenRun_UsesReplacementModule()
    {
        // Arrange
        var rangeCommands = _commands;
        const string moduleName = "ReopenReplaceModule";
        string originalCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""original-module""
End Sub";
        string replacementCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""replacement-run-ok""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, originalCode);
        await _fixture.SaveAndReopenAsync();

        // Act
        var reopenedBatch = _fixture.BatchToken;
        Delete(reopenedBatch, moduleName);
        var afterDeleteList = _scriptCommands.List(reopenedBatch);
        Import(reopenedBatch, moduleName, replacementCode);
        var replacementView = _scriptCommands.View(reopenedBatch, moduleName);
        _scriptCommands.Run(reopenedBatch, $"{moduleName}.WriteMarker", null);

        // Assert
        Assert.True(afterDeleteList.Success, $"List should succeed after delete. Error: {afterDeleteList.ErrorMessage}");
        Assert.DoesNotContain(afterDeleteList.Scripts, s => s.Name == moduleName);

        Assert.True(replacementView.Success, $"View should succeed after replacement import. Error: {replacementView.ErrorMessage}");
        Assert.Contains("replacement-run-ok", replacementView.Code);
        Assert.DoesNotContain("original-module", replacementView.Code);

        var cellResult = rangeCommands.GetValues(reopenedBatch, "Sheet1", "A1");
        Assert.True(cellResult.Success, $"GetValues should succeed after replacement run. Error: {cellResult.ErrorMessage}");
        Assert.NotNull(cellResult.Values);
        Assert.Single(cellResult.Values);
        Assert.Equal("replacement-run-ok", cellResult.Values[0][0]?.ToString());
    }

    [Fact]
    public void ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // Import a module first
        string vbaCode = "Sub TestCode()\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

        // Act - Delete the module
        Delete(batch, "TestModule");

        // Verify module is gone
        var listResult = _scriptCommands.List(batch);
        Assert.DoesNotContain(listResult.Scripts, s => s.Name == "TestModule");
    }
    [Fact]
    public void ScriptCommands_View_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // Import a module with known code
        string expectedCode = "Sub ViewTest()\n    MsgBox \"Hello\"\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "ViewTestModule", expectedCode);

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
        // Import initial module
        string initialCode = "Sub OriginalCode()\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "UpdateTestModule", initialCode);

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

    [Fact]
    public void ScriptCommands_Run_WithParameters_PassesArgumentsCorrectly()
    {
        // Arrange
        // Import a macro that writes a parameter to a cell for verification
        string vbaCode = @"Sub TestWithParam(value As String)
    ThisWorkbook.Sheets(1).Range(""A1"").Value = value
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "ParamTest", vbaCode);

        // Act - Run with parameter
        _scriptCommands.Run(batch, "ParamTest.TestWithParam", null, "HelloWorld");

        // Assert - Verify the macro wrote the value
        var rangeCommands = _commands;
        var result = rangeCommands.GetValues(batch, "Sheet1", "A1");
        Assert.True(result.Success, $"GetValues should succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Values);
        Assert.Single(result.Values);
        Assert.Equal("HelloWorld", result.Values[0][0]?.ToString());
    }

    [Fact]
    public void ScriptCommands_Run_WithMultipleParameters_PassesAllArguments()
    {
        // Arrange
        // Import a macro that writes two parameters to separate cells
        string vbaCode = @"Sub TestMultiParam(val1 As String, val2 As String)
    ThisWorkbook.Sheets(1).Range(""A1"").Value = val1
    ThisWorkbook.Sheets(1).Range(""B1"").Value = val2
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "MultiParamTest", vbaCode);

        // Act - Run with two parameters
        _scriptCommands.Run(batch, "MultiParamTest.TestMultiParam", null, "First", "Second");

        // Assert - Verify both parameters were passed correctly
        var rangeCommands = _commands;
        var result = rangeCommands.GetValues(batch, "Sheet1", "A1:B1");
        Assert.True(result.Success, $"GetValues should succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Values);
        Assert.Single(result.Values); // one row
        Assert.Equal("First", result.Values[0][0]?.ToString());
        Assert.Equal("Second", result.Values[0][1]?.ToString());
    }

    private void Import(
        IExcelBatch batch,
        string moduleName,
        string vbaCode)
    {
        _scriptCommands.Import(batch, moduleName, vbaCode);
        _fixture.RegisterVbaModuleForCleanup(moduleName);
    }

    private void Delete(IExcelBatch batch, string moduleName)
    {
        _scriptCommands.Delete(batch, moduleName);
        _fixture.ForgetVbaModule(moduleName);
    }
}

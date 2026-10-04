using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for VBA operations on desktop Excel with VBA project access enabled.
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
    public void ScriptCommands_Import_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        const string vbaCode = "Sub TestImport()\n    MsgBox \"Hello\"\nEnd Sub";

        // Act
        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

        // Assert - verify module exists via list
        var listResult = RequireSuccess(_scriptCommands.List(batch));
        Assert.Single(listResult.Scripts, s => s.Name == "TestModule");
        var code = RequireSuccess(_scriptCommands.View(batch, "TestModule"));
        AssertStoredCode(vbaCode, code.Code);
    }
    [Fact]
    public void ScriptCommands_Run_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // Import a test macro first
        string vbaCode = @"Sub TestProcedure()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""macro-ran""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

        // Act - Run the macro
        RequireSuccess(_scriptCommands.Run(batch, "TestModule.TestProcedure", null));

        // Assert - No exception thrown; to be thorough, ensure module still exists
        var listResult = _scriptCommands.List(batch);
        Assert.True(listResult.Success, listResult.ErrorMessage);
        Assert.Contains(listResult.Scripts, s => s.Name == "TestModule");
        var values = _commands.GetValues(batch, "Sheet1", "A1");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("macro-ran", Assert.Single(Assert.Single(values.Values)));
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

        RequireSuccess(_scriptCommands.Run(reopenedBatch, "ReopenTestModule.WriteMarker", null));

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
        RequireSuccess(_scriptCommands.Update(reopenedBatch, moduleName, updatedCode));
        var viewResult = _scriptCommands.View(reopenedBatch, moduleName);
        RequireSuccess(_scriptCommands.Run(reopenedBatch, $"{moduleName}.WriteMarker", null));

        // Assert
        Assert.True(viewResult.Success, $"View should succeed after reopened update. Error: {viewResult.ErrorMessage}");
        Assert.Contains("updated-run-ok", viewResult.Code);
        Assert.DoesNotContain("original-run", viewResult.Code);
        AssertStoredCode(updatedCode, viewResult.Code);

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
        RequireSuccess(_scriptCommands.Run(reopenedBatch, $"{moduleName}.WriteMarker", null));

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
        Import(batch, "RetainedModule", "Sub Retained()\nEnd Sub");
        var retained = RequireSuccess(_scriptCommands.View(batch, "RetainedModule")).Code;

        // Act - Delete the module
        Delete(batch, "TestModule");

        // Verify module is gone
        var listResult = RequireSuccess(_scriptCommands.List(batch));
        Assert.DoesNotContain(listResult.Scripts, s => s.Name == "TestModule");
        Assert.Single(listResult.Scripts, s => s.Name == "RetainedModule");
        Assert.Equal(retained, RequireSuccess(_scriptCommands.View(batch, "RetainedModule")).Code);
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
        RequireSuccess(_scriptCommands.Update(batch, "UpdateTestModule", updatedCode));

        // Verify the code was updated
        var viewResult = _scriptCommands.View(batch, "UpdateTestModule");
        Assert.True(viewResult.Success);
        Assert.Contains("UpdatedCode", viewResult.Code);
        Assert.Contains("Updated", viewResult.Code);
        Assert.DoesNotContain("OriginalCode", viewResult.Code);
        AssertStoredCode(updatedCode, viewResult.Code);
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
        RequireSuccess(_scriptCommands.Run(batch, "ParamTest.TestWithParam", null, "HelloWorld"));

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
        RequireSuccess(_scriptCommands.Run(batch, "MultiParamTest.TestMultiParam", null, "First", "Second"));

        // Assert - Verify both parameters were passed correctly
        var rangeCommands = _commands;
        var result = rangeCommands.GetValues(batch, "Sheet1", "A1:B1");
        Assert.True(result.Success, $"GetValues should succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Values);
        Assert.Single(result.Values); // one row
        Assert.Equal("First", result.Values[0][0]?.ToString());
        Assert.Equal("Second", result.Values[0][1]?.ToString());
    }

    private static void AssertStoredCode(string expected, string actual)
    {
        expected = expected.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace("\n", "\r\n", StringComparison.Ordinal);
        actual = actual.Trim();
        // VBE canonicalizes identifier casing; literals and comments must remain exact.
        Assert.Equal(expected, actual, ignoreCase: true);
        const string literalsAndComments = "\"(?:[^\"\\r\\n]|\"\")*\"|'[^\\r\\n]*";
        Assert.Equal(
            System.Text.RegularExpressions.Regex.Matches(expected, literalsAndComments).Select(match => match.Value),
            System.Text.RegularExpressions.Regex.Matches(actual, literalsAndComments).Select(match => match.Value));
    }

    private void Import(
        IExcelBatch batch,
        string moduleName,
        string vbaCode)
    {
        RequireSuccess(_scriptCommands.Import(batch, moduleName, vbaCode));
        _fixture.RegisterVbaModuleForCleanup(moduleName);
    }

    private void Delete(IExcelBatch batch, string moduleName)
    {
        RequireSuccess(_scriptCommands.Delete(batch, moduleName));
        _fixture.ForgetVbaModule(moduleName);
    }
}

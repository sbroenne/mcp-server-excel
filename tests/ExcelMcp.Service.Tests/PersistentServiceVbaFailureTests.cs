using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class PersistentServiceVbaFailureTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceVbaFixture>
{
    private readonly IPersistentVbaCommands _vba;
    private string _guardSheetName = string.Empty;
    private string _guardCode = string.Empty;
    private string _guardModules = string.Empty;

    public PersistentServiceVbaFailureTests(
        PersistentServiceVbaFixture fixture) :
        base(fixture)
    {
        _vba = fixture.CreateCommands<IPersistentVbaCommands>();
    }

    [Fact]
    public async Task Import_ExistingModuleName_HasConflictCategory()
    {
        const string moduleName = "DuplicateModule";
        const string vbaCode = "Sub Test()\nEnd Sub";
        RequireSuccess(_vba.Import(_fixture.BatchToken, moduleName, vbaCode));
        _fixture.RegisterVbaModuleForCleanup(moduleName);
        SeedGuard();

        var response = await _fixture.SendForFailureAsync(
            "vba.import",
            new { moduleName, vbaCode });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.Conflict.ToString(), response.ErrorCategory);
        Assert.Contains("already exists", response.ErrorMessage);
        var retained = _vba.View(_fixture.BatchToken, moduleName);
        Assert.True(retained.Success, retained.ErrorMessage);
        Assert.Equal(vbaCode.Replace("\n", "\r\n", StringComparison.Ordinal), retained.Code.Trim());
        var modules = _vba.List(_fixture.BatchToken);
        Assert.True(modules.Success, modules.ErrorMessage);
        Assert.Single(modules.Scripts, module => module.Name == moduleName);
        AssertGuardPreserved();
    }

    [Fact]
    public async Task Delete_MissingModule_HasNotFoundCategory()
    {
        SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "vba.delete",
            new { moduleName = "NonExistentModule" });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.NotFound.ToString(), response.ErrorCategory);
        Assert.Contains("not found", response.ErrorMessage);
        AssertGuardPreserved();
    }

    [Fact]
    public async Task View_MissingModule_HasNotFoundCategory()
    {
        SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "vba.view",
            new { moduleName = "NonExistentModule" });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.NotFound.ToString(), response.ErrorCategory);
        Assert.Contains("not found", response.ErrorMessage);
        AssertGuardPreserved();
    }

    [Fact]
    public async Task Run_MissingProcedure_RemainsComInteropFailure()
    {
        SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "vba.run",
            new
            {
                procedureName = "NonExistentModule.NonExistentProc",
                timeout = (int?)null,
                parameters = Array.Empty<string>()
            });

        Assert.Equal("ComInterop", response.ErrorCategory);
        Assert.Contains(
            "nonexistent",
            response.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
        AssertGuardPreserved();
    }

    [Fact]
    public async Task Run_EmptyProcedureName_HasInvalidInputCategory()
    {
        SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "vba.run",
            new
            {
                procedureName = string.Empty,
                timeout = (int?)null,
                parameters = Array.Empty<string>()
            });

        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
        AssertGuardPreserved();
    }

    [Fact]
    public async Task Import_EmptyModuleName_HasInvalidInputCategory()
    {
        SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "vba.import",
            new
            {
                moduleName = string.Empty,
                vbaCode = "Sub Test()\nEnd Sub"
            });

        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
        AssertGuardPreserved();
    }

    private void SeedGuard()
    {
        _guardSheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, _guardSheetName, "A1:B1", [["preserved", 11]]));
        var code = $"Sub RetainedProcedure()\n    ThisWorkbook.Sheets(\"{_guardSheetName}\").Range(\"B1\").Value = 19\nEnd Sub";
        RequireSuccess(_vba.Import(_fixture.BatchToken, "RetainedModule", code));
        _fixture.RegisterVbaModuleForCleanup("RetainedModule");
        _guardCode = RequireSuccess(_vba.View(_fixture.BatchToken, "RetainedModule")).Code;
        _guardModules = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_vba.List(_fixture.BatchToken)).Scripts);
    }

    private void AssertGuardPreserved()
    {
        Assert.Equal(_guardCode, RequireSuccess(_vba.View(_fixture.BatchToken, "RetainedModule")).Code);
        Assert.Equal(_guardModules, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_vba.List(_fixture.BatchToken)).Scripts));
        Assert.Equal(["preserved", 11], Assert.Single(
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, _guardSheetName, "A1:B1")).Values));
        RequireSuccess(_vba.Run(_fixture.BatchToken, "RetainedModule.RetainedProcedure", null));
        Assert.Equal(["preserved", 19], Assert.Single(
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, _guardSheetName, "A1:B1")).Values));
    }
}

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class PersistentServiceVbaUnsupportedFormatTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceWorkbookFixture>
{
    public PersistentServiceVbaUnsupportedFormatTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
    }

    [Fact]
    public async Task Import_UnsupportedFormat_HasInvalidInputCategory()
    {
        var sheetName = SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "vba.import",
            new
            {
                moduleName = "TestModule",
                vbaCode = "Sub Test()\nEnd Sub"
            });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
        Assert.Contains("macro-enabled", response.ErrorMessage);
        AssertGuardPreserved(sheetName);
    }

    [Fact]
    public async Task List_UnsupportedFormat_HasInvalidInputCategory()
    {
        var sheetName = SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "vba.list",
            new { });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
        Assert.Contains("macro-enabled", response.ErrorMessage);
        AssertGuardPreserved(sheetName);
    }

    private string SeedGuard()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B1", [["preserved", 11]]));
        return sheetName;
    }

    private void AssertGuardPreserved(string sheetName)
    {
        Assert.Equal(["preserved", 11], Assert.Single(
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:B1")).Values));
        _fixture.ExecuteRawVerification((context, _) =>
            Assert.Equal(Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, context.Book.FileFormat));
    }
}

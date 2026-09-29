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
        _vba.Import(_fixture.BatchToken, moduleName, vbaCode);
        _fixture.RegisterVbaModuleForCleanup(moduleName);

        var response = await _fixture.SendForFailureAsync(
            "vba.import",
            new { moduleName, vbaCode });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.Conflict.ToString(), response.ErrorCategory);
        Assert.Contains("already exists", response.ErrorMessage);
    }

    [Fact]
    public async Task Delete_MissingModule_HasNotFoundCategory()
    {
        var response = await _fixture.SendForFailureAsync(
            "vba.delete",
            new { moduleName = "NonExistentModule" });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.NotFound.ToString(), response.ErrorCategory);
        Assert.Contains("not found", response.ErrorMessage);
    }

    [Fact]
    public async Task View_MissingModule_HasNotFoundCategory()
    {
        var response = await _fixture.SendForFailureAsync(
            "vba.view",
            new { moduleName = "NonExistentModule" });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.NotFound.ToString(), response.ErrorCategory);
        Assert.Contains("not found", response.ErrorMessage);
    }

    [Fact]
    public async Task Run_MissingProcedure_RemainsComInteropFailure()
    {
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
    }

    [Fact]
    public async Task Run_EmptyProcedureName_HasInvalidInputCategory()
    {
        var response = await _fixture.SendForFailureAsync(
            "vba.run",
            new
            {
                procedureName = string.Empty,
                timeout = (int?)null,
                parameters = Array.Empty<string>()
            });

        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
    }

    [Fact]
    public async Task Import_EmptyModuleName_HasInvalidInputCategory()
    {
        var response = await _fixture.SendForFailureAsync(
            "vba.import",
            new
            {
                moduleName = string.Empty,
                vbaCode = "Sub Test()\nEnd Sub"
            });

        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
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
    }

    [Fact]
    public async Task List_UnsupportedFormat_HasInvalidInputCategory()
    {
        var response = await _fixture.SendForFailureAsync(
            "vba.list",
            new { });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
        Assert.Contains("macro-enabled", response.ErrorMessage);
    }
}

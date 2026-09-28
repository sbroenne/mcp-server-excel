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
    public void Import_ExistingModuleName_ThrowsInvalidOperationException()
    {
        const string moduleName = "DuplicateModule";
        const string vbaCode = "Sub Test()\nEnd Sub";
        _vba.Import(_fixture.BatchToken, moduleName, vbaCode);
        _fixture.RegisterVbaModuleForCleanup(moduleName);

        var exception = Assert.Throws<InvalidOperationException>(
            () => _vba.Import(_fixture.BatchToken, moduleName, vbaCode));

        Assert.Contains("already exists", exception.Message);
    }

    [Fact]
    public void Delete_MissingModule_ThrowsInvalidOperationException()
    {
        var exception = Assert.Throws<InvalidOperationException>(
            () => _vba.Delete(_fixture.BatchToken, "NonExistentModule"));

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void View_MissingModule_ThrowsInvalidOperationException()
    {
        var exception = Assert.Throws<InvalidOperationException>(
            () => _vba.View(_fixture.BatchToken, "NonExistentModule"));

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Run_MissingProcedure_ThrowsComException()
    {
        var exception = Assert.Throws<InvalidOperationException>(
            () => _vba.Run(
                _fixture.BatchToken,
                "NonExistentModule.NonExistentProc",
                null));

        Assert.Contains("COMException", exception.Message);
        Assert.Contains(
            "nonexistent",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Run_EmptyProcedureName_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(
            () => _vba.Run(_fixture.BatchToken, string.Empty, null));
    }

    [Fact]
    public void Import_EmptyModuleName_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(
            () => _vba.Import(
                _fixture.BatchToken,
                string.Empty,
                "Sub Test()\nEnd Sub"));
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
    private readonly IPersistentVbaCommands _vba;

    public PersistentServiceVbaUnsupportedFormatTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _vba = fixture.CreateCommands<IPersistentVbaCommands>();
    }

    [Fact]
    public void Import_UnsupportedFormat_ThrowsInvalidOperationException()
    {
        var exception = Assert.Throws<InvalidOperationException>(
            () => _vba.Import(
                _fixture.BatchToken,
                "TestModule",
                "Sub Test()\nEnd Sub"));

        Assert.Contains("macro-enabled", exception.Message);
    }
}

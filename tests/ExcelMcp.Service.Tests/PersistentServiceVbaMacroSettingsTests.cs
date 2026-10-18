using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Reopened macro workbooks follow the user's Excel macro settings. The fixture
/// workbook lives in an untrusted temporary folder, so with Excel's default
/// settings its macros are blocked after reopening. These tests require that
/// default: no "enable all macros" setting and no trusted location for Temp.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class PersistentServiceVbaMacroSettingsTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceReopenedVbaFixture>
{
    private readonly IPersistentVbaCommands _vba;

    public PersistentServiceVbaMacroSettingsTests(
        PersistentServiceReopenedVbaFixture fixture) :
        base(fixture)
    {
        _vba = fixture.CreateCommands<IPersistentVbaCommands>();
    }

    [Fact]
    public async Task ReopenedMacroWorkbook_DoesNotRunWorkbookOpen()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_vba.List(_fixture.BatchToken));
        var documentModule = _fixture.ExecuteRawVerification((context, _) => (string)context.Book.CodeName);
        Assert.False(string.IsNullOrWhiteSpace(documentModule));
        var originalCode = RequireSuccess(_vba.View(_fixture.BatchToken, documentModule)).Code;
        var openCode =
            "Private Sub Workbook_Open()\n" +
            $"    ThisWorkbook.Sheets(\"{sheetName}\").Range(\"A1\").Value = \"open-ran\"\n" +
            "End Sub";

        try
        {
            RequireSuccess(_vba.Update(_fixture.BatchToken, documentModule, openCode));
            await _fixture.SaveAndReopenAsync();

            Assert.Contains(
                "open-ran",
                RequireSuccess(_vba.View(_fixture.BatchToken, documentModule)).Code,
                StringComparison.Ordinal);
            var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1"));
            Assert.Null(Assert.Single(Assert.Single(values.Values)));
        }
        finally
        {
            RequireSuccess(_vba.Update(
                _fixture.BatchToken,
                documentModule,
                string.IsNullOrWhiteSpace(originalCode) ? "Option Explicit" : originalCode));
        }
    }

    [Fact]
    public async Task Run_OnReopenedMacroWorkbook_ReportsBlockedMacrosWithoutRunning()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        const string moduleName = "BlockedRunModule";
        Import(moduleName, MarkerProcedure(sheetName, "reopened-run"));
        await _fixture.SaveAndReopenAsync();

        var modules = RequireSuccess(_vba.List(_fixture.BatchToken));
        Assert.Single(modules.Scripts, module => module.Name == moduleName);

        await AssertRunBlockedAsync($"{moduleName}.WriteMarker", sheetName);
    }

    [Fact]
    public async Task Update_OnReopenedMacroWorkbook_ChangesCodeButRunIsBlocked()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        const string moduleName = "BlockedUpdateModule";
        Import(moduleName, MarkerProcedure(sheetName, "original-run"));
        await _fixture.SaveAndReopenAsync();

        var updatedCode = MarkerProcedure(sheetName, "updated-run");
        RequireSuccess(_vba.Update(_fixture.BatchToken, moduleName, updatedCode));

        var view = RequireSuccess(_vba.View(_fixture.BatchToken, moduleName));
        Assert.Contains("updated-run", view.Code, StringComparison.Ordinal);
        Assert.DoesNotContain("original-run", view.Code, StringComparison.Ordinal);
        await AssertRunBlockedAsync($"{moduleName}.WriteMarker", sheetName);
    }

    [Fact]
    public async Task DeleteThenImport_OnReopenedMacroWorkbook_ReplacesModuleButRunIsBlocked()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        const string moduleName = "BlockedReplaceModule";
        Import(moduleName, MarkerProcedure(sheetName, "original-module"));
        await _fixture.SaveAndReopenAsync();

        RequireSuccess(_vba.Delete(_fixture.BatchToken, moduleName));
        _fixture.ForgetVbaModule(moduleName);
        var afterDelete = RequireSuccess(_vba.List(_fixture.BatchToken));
        Assert.DoesNotContain(afterDelete.Scripts, module => module.Name == moduleName);

        Import(moduleName, MarkerProcedure(sheetName, "replacement-module"));
        var view = RequireSuccess(_vba.View(_fixture.BatchToken, moduleName));
        Assert.Contains("replacement-module", view.Code, StringComparison.Ordinal);
        Assert.DoesNotContain("original-module", view.Code, StringComparison.Ordinal);
        await AssertRunBlockedAsync($"{moduleName}.WriteMarker", sheetName);
    }

    private async Task AssertRunBlockedAsync(string procedureName, string sheetName)
    {
        var response = await _fixture.SendForFailureAsync(
            "vba.run",
            new
            {
                procedureName,
                timeout = (int?)null,
                parameters = Array.Empty<string>()
            });

        Assert.Equal("ComInterop", response.ErrorCategory);
        Assert.Contains(procedureName, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("Excel macro settings", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("trusted location", response.ErrorMessage, StringComparison.Ordinal);
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1"));
        Assert.Null(Assert.Single(Assert.Single(values.Values)));
    }

    private static string MarkerProcedure(string sheetName, string marker) =>
        "Sub WriteMarker()\n" +
        $"    ThisWorkbook.Sheets(\"{sheetName}\").Range(\"A1\").Value = \"{marker}\"\n" +
        "End Sub";

    private void Import(string moduleName, string vbaCode)
    {
        RequireSuccess(_vba.Import(_fixture.BatchToken, moduleName, vbaCode));
        _fixture.RegisterVbaModuleForCleanup(moduleName);
    }
}

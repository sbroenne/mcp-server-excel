using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Workbook;

/// <summary>
/// Integration tests for workbook protection operations.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Workbook")]
public class WorkbookCommandsTests : IClassFixture<SheetTestsFixture>
{
    private readonly WorkbookCommands _workbookCommands;
    private readonly SheetTestsFixture _fixture;

    public WorkbookCommandsTests(SheetTestsFixture fixture)
    {
        _workbookCommands = new WorkbookCommands();
        _fixture = fixture;
    }

    [Fact]
    public void SetProtection_ProtectsAndUnprotectsWorkbook()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        var protectResult = _workbookCommands.SetProtection(batch, true);
        Assert.True(protectResult.Success, $"Expected protect to succeed but got error: {protectResult.ErrorMessage}");
        Assert.True(IsWorkbookProtected(batch));

        var unprotectResult = _workbookCommands.SetProtection(batch, false);
        Assert.True(unprotectResult.Success, $"Expected unprotect to succeed but got error: {unprotectResult.ErrorMessage}");
        Assert.False(IsWorkbookProtected(batch));
    }

    [Fact]
    public void SetViewOptions_UpdatesGridlinesAndHeadings()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        var setResult = _workbookCommands.SetViewOptions(batch, displayGridlines: false, displayHeadings: true);
        Assert.True(setResult.Success, $"Expected view options to update but got error: {setResult.ErrorMessage}");

        var getResult = _workbookCommands.GetViewOptions(batch);
        Assert.True(getResult.Success, $"Expected view options to be read back but got error: {getResult.ErrorMessage}");
        Assert.False(getResult.DisplayGridlines);
        Assert.True(getResult.DisplayHeadings);
    }

    private static bool IsWorkbookProtected(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            return ctx.Book.ProtectStructure || ctx.Book.ProtectWindows;
        });
    }
}

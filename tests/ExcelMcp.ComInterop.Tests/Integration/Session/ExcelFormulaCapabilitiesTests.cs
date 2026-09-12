using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.ComInterop.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelContext")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")]
public sealed class ExcelFormulaCapabilitiesTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;

    public ExcelFormulaCapabilitiesTests(TempDirectoryFixture fixture) => _fixture = fixture;

    [Fact]
    public void Capabilities_ReadProbe_PreservesProtectedCellAndWorkbookState()
    {
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(_fixture.CreateFilePath(), show: false);
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);

        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                cell = sheet.Range["A1"];
                cell.Formula = "=1+2";
                sheet.Protect();
                bool saved = ctx.Book.Saved;
                var calculation = ctx.App.Calculation;
                bool eventsEnabled = ctx.App.EnableEvents;

                var capabilities = ctx.Capabilities;
                bool supported = capabilities.SupportsFormula2;
                Assert.Equal(supported, capabilities.SupportsFormula2);

                Assert.Equal("=1+2", cell.Formula);
                Assert.Equal(3.0, Convert.ToDouble(cell.Value2));
                Assert.True(sheet.ProtectContents);
                Assert.Equal(saved, ctx.Book.Saved);
                Assert.Equal(calculation, ctx.App.Calculation);
                Assert.Equal(eventsEnabled, ctx.App.EnableEvents);
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public void Capabilities_PathUpdate_PreservesSessionCache()
    {
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(_fixture.CreateFilePath(), show: false);
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);
        var capabilities = batch.Execute((ctx, ct) =>
        {
            _ = ctx.Capabilities.SupportsFormula2;
            return ctx.Capabilities;
        });

        string targetPath = _fixture.CreateFilePath();
        batch.Execute((ctx, ct) => ctx.Book.SaveAs(targetPath));
        batch.UpdateWorkbookPath(targetPath);

        batch.Execute((ctx, ct) =>
        {
            Assert.Equal(targetPath, ctx.WorkbookPath);
            Assert.Same(capabilities, ctx.Capabilities);
            Assert.Equal(capabilities.SupportsFormula2, ctx.Capabilities.SupportsFormula2);
        });
    }
}

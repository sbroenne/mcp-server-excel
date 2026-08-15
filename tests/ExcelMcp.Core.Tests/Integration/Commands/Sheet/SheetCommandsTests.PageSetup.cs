using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Tests for worksheet page setup operations.
/// </summary>
public partial class SheetCommandsTests
{
    [Fact]
    public void SetPageSetup_UpdatesSheetPageSetupProperties()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"PageSetup_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        var setResult = _sheetCommands.SetPageSetup(batch, sheetName, "landscape", 1, 2, false, true);
        Assert.True(setResult.Success, $"Expected page setup to succeed but got error: {setResult.ErrorMessage}");

        var getResult = _sheetCommands.GetPageSetup(batch, sheetName);
        Assert.True(getResult.Success, $"Expected page setup read to succeed but got error: {getResult.ErrorMessage}");
        Assert.Equal("landscape", getResult.Orientation);
        Assert.Equal(1, getResult.FitToPagesWide);
        Assert.Equal(2, getResult.FitToPagesTall);
        Assert.False(getResult.CenterHorizontally);
        Assert.True(getResult.CenterVertically);
        Assert.False(IsPageSetupZoomEnabled(batch, sheetName));
    }

    [Fact]
    public void GetPageSetup_AutomaticScaling_ReturnsNullFitValues()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"AutoScale_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        var result = _sheetCommands.GetPageSetup(batch, sheetName);

        Assert.True(result.Success, $"Expected page setup read to succeed but got error: {result.ErrorMessage}");
        Assert.Null(result.FitToPagesWide);
        Assert.Null(result.FitToPagesTall);
    }

    private static bool IsPageSetupZoomEnabled(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PageSetup? pageSetup = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                pageSetup = sheet.PageSetup;
                return pageSetup.Zoom is not bool value || value;
            }
            finally
            {
                ComUtilities.Release(ref pageSetup);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

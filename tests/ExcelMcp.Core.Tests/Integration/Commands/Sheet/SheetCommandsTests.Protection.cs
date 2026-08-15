using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Tests for worksheet protection operations.
/// </summary>
public partial class SheetCommandsTests
{
    [Fact]
    public void SetProtection_ProtectsAndUnprotectsSheet()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Protect_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        var protectResult = _sheetCommands.SetProtection(batch, sheetName, true);
        Assert.True(protectResult.Success, $"Expected protect to succeed but got error: {protectResult.ErrorMessage}");
        Assert.True(IsSheetProtected(batch, sheetName));

        var unprotectResult = _sheetCommands.SetProtection(batch, sheetName, false);
        Assert.True(unprotectResult.Success, $"Expected unprotect to succeed but got error: {unprotectResult.ErrorMessage}");
        Assert.False(IsSheetProtected(batch, sheetName));
    }

    private static bool IsSheetProtected(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                return sheet.ProtectContents;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

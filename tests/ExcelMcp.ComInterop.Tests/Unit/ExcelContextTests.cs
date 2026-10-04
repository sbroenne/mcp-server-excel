using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

/// <summary>
/// Tests constructor validation order without creating Excel COM objects.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("RequiresExcel", "false")]
public class ExcelContextTests
{
    [Fact]
    public void Constructor_WithNullWorkbookPath_ThrowsBeforeNullExcel()
    {
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(null!, null!, null!));

        Assert.Equal("workbookPath", ex.ParamName);
    }

    [Theory]
    [InlineData(@"C:\test\workbook.xlsx")]
    [InlineData(@"\\server\share\workbook.xlsm")]
    [InlineData(@"D:\Documents\My Workbook.xlsx")]
    [InlineData(@"workbook.xlsx")] // Relative path
    public void Constructor_WithNullExcelAnyPath_ThrowsArgumentNullException(string workbookPath)
    {
        // Act & Assert - Path is validated, then excel COM object is validated
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath, null!, null!));

        // excel is the first COM parameter validated after workbookPath
        Assert.Equal("excel", ex.ParamName);
    }
}




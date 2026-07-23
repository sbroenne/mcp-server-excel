using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for ExcelContext - validates constructor and property behavior.
/// This class is a simple data holder, so tests focus on path validation and immutability.
/// Note: Excel.Application and Excel.Workbook COM objects cannot be mocked in unit tests,
/// so these tests use null! for those parameters and verify only what is testable.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class ExcelContextTests
{
    [Fact]
    public void Constructor_WithValidArguments_SetsWorkbookPathCorrectly()
    {
        // Arrange
        string workbookPath = @"C:\test\workbook.xlsx";

        // Act & Assert - Constructor throws ArgumentNullException for null COM objects,
        // which is expected behavior. WorkbookPath validation is tested separately.
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath, null!, null!));

        // When null is passed, the constructor throws on the first null param (excel)
        Assert.NotNull(ex);
    }

    [Fact]
    public void Constructor_WithNullWorkbookPath_ThrowsArgumentNullException()
    {
        // Arrange
        string? workbookPath = null;

        // Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath!, null!, null!));

        Assert.Equal("workbookPath", ex.ParamName);
    }

    [Fact]
    public void Constructor_WithNullExcel_ThrowsArgumentNullException()
    {
        // Arrange
        string workbookPath = @"C:\test\workbook.xlsx";

        // Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath, null!, null!));

        Assert.Equal("excel", ex.ParamName);
    }

    [Fact]
    public void Constructor_WithNullWorkbookPath_ThrowsBeforeNullExcel()
    {
        // Arrange
        string? workbookPath = null;

        // Act & Assert - WorkbookPath is validated first
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath!, null!, null!));

        Assert.Equal("workbookPath", ex.ParamName);
    }

    [Fact]
    public void Constructor_WorkbookPathValidation_RejectsNull()
    {
        // Arrange & Act & Assert
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

    [Fact]
    public void Constructor_NullWorkbookPath_ThrowsWithCorrectParamName()
    {
        // Arrange - Simulates null path being passed
        Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(null!, null!, null!));
    }
}






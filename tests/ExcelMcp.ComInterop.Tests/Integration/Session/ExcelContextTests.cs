using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Unit tests for ExcelContext - validates constructor and property behavior.
/// This class is a simple data holder, so tests focus on validation and immutability.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class ExcelContextTests
{
    [Fact]
    public void Constructor_WithValidArguments_SetsPropertiesCorrectly()
    {
        // Arrange
        string workbookPath = @"C:\test\workbook.xlsx";
        var mockExcel = new { Name = "Excel.Application" };
        var mockWorkbook = new { Name = "Workbook1" };

        // Act
        var context = new ExcelContext(workbookPath, mockExcel, mockWorkbook);

        // Assert
        Assert.Equal(workbookPath, context.WorkbookPath);
        Assert.Same(mockExcel, context.App);
        Assert.Same(mockWorkbook, context.Book);
    }

    [Fact]
    public void Constructor_WithNullWorkbookPath_ThrowsArgumentNullException()
    {
        // Arrange
        string? workbookPath = null;
        var mockExcel = new { Name = "Excel.Application" };
        var mockWorkbook = new { Name = "Workbook1" };

        // Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath!, mockExcel, mockWorkbook));

        Assert.Equal("workbookPath", ex.ParamName);
    }

    [Fact]
    public void Constructor_WithNullExcel_ThrowsArgumentNullException()
    {
        // Arrange
        string workbookPath = @"C:\test\workbook.xlsx";
        dynamic? mockExcel = null;
        var mockWorkbook = new { Name = "Workbook1" };

        // Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath, mockExcel!, mockWorkbook));

        Assert.Equal("excel", ex.ParamName);
    }

    [Fact]
    public void Constructor_WithNullWorkbook_ThrowsArgumentNullException()
    {
        // Arrange
        string workbookPath = @"C:\test\workbook.xlsx";
        var mockExcel = new { Name = "Excel.Application" };
        dynamic? mockWorkbook = null;

        // Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new ExcelContext(workbookPath, mockExcel, mockWorkbook!));

        Assert.Equal("workbook", ex.ParamName);
    }

    [Fact]
    public void Properties_AreReadOnly_CannotBeModified()
    {
        // Arrange
        string workbookPath = @"C:\test\workbook.xlsx";
        var mockExcel = new { Name = "Excel.Application" };
        var mockWorkbook = new { Name = "Workbook1" };
        var context = new ExcelContext(workbookPath, mockExcel, mockWorkbook);

        // Act & Assert - Verify properties are get-only
        // This is a compile-time check, but we can verify the values don't change
        Assert.Equal(workbookPath, context.WorkbookPath);
        Assert.Same(mockExcel, context.App);
        Assert.Same(mockWorkbook, context.Book);

        // Create another context and verify independence
        var context2 = new ExcelContext(@"C:\other.xlsx", mockExcel, mockWorkbook);
        Assert.NotEqual(context.WorkbookPath, context2.WorkbookPath);
        Assert.Same(mockExcel, context2.App); // Same references
        Assert.Same(mockWorkbook, context2.Book);
    }

    [Fact]
    public void Constructor_WithEmptyWorkbookPath_AllowsEmptyString()
    {
        // Arrange
        string workbookPath = string.Empty;
        var mockExcel = new { Name = "Excel.Application" };
        var mockWorkbook = new { Name = "Workbook1" };

        // Act
        var context = new ExcelContext(workbookPath, mockExcel, mockWorkbook);

        // Assert - Empty string is allowed (validation happens at higher levels)
        Assert.Equal(string.Empty, context.WorkbookPath);
    }

    [Theory]
    [InlineData(@"C:\test\workbook.xlsx")]
    [InlineData(@"\\server\share\workbook.xlsm")]
    [InlineData(@"D:\Documents\My Workbook.xlsx")]
    [InlineData(@"workbook.xlsx")] // Relative path
    public void Constructor_WithVariousPathFormats_StoresPathAsProvided(string workbookPath)
    {
        // Arrange
        var mockExcel = new { Name = "Excel.Application" };
        var mockWorkbook = new { Name = "Workbook1" };

        // Act
        var context = new ExcelContext(workbookPath, mockExcel, mockWorkbook);

        // Assert - Path is stored exactly as provided (normalization happens elsewhere)
        Assert.Equal(workbookPath, context.WorkbookPath);
    }

    [Fact]
    public void MultipleContexts_WithSameComObjects_CanCoexist()
    {
        // Arrange - Simulates sharing COM objects between contexts
        var sharedExcel = new { Name = "Excel.Application" };
        var sharedWorkbook = new { Name = "Workbook1" };

        // Act
        var context1 = new ExcelContext(@"C:\test1.xlsx", sharedExcel, sharedWorkbook);
        var context2 = new ExcelContext(@"C:\test2.xlsx", sharedExcel, sharedWorkbook);

        // Assert - Both contexts can exist with same COM references
        Assert.NotSame(context1, context2);
        Assert.Same(context1.App, context2.App);
        Assert.Same(context1.Book, context2.Book);
        Assert.NotEqual(context1.WorkbookPath, context2.WorkbookPath);
    }
}

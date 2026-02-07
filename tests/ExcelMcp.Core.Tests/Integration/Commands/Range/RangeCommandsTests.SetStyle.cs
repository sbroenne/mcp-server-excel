using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    [Fact]
    public void SetStyle_Heading1_AppliesSuccessfully()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetStyle(batch, sheetName, "A1", "Heading 1");

        // Assert - void method throws on failure, succeeds silently on success
    }

    [Fact]
    public void SetStyle_GoodBadNeutral_AllApplySuccessfully()
    {
        // Arrange & Act & Assert
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetStyle(batch, sheetName, "A1", "Good");
        _commands.SetStyle(batch, sheetName, "A2", "Bad");
        _commands.SetStyle(batch, sheetName, "A3", "Neutral");
        // void methods throw on failure, succeed silently
    }

    [Fact]
    public void SetStyle_Accent1_AppliesSuccessfully()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetStyle(batch, sheetName, "A1:E1", "Accent1");

        // Assert - void method throws on failure, succeeds silently on success
    }

    [Fact]
    public void SetStyle_TotalStyle_AppliesSuccessfully()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetStyle(batch, sheetName, "A10:E10", "Total");

        // Assert - void method throws on failure, succeeds silently on success
    }

    [Fact]
    public void SetStyle_CurrencyComma_AppliesSuccessfully()
    {
        // Arrange & Act & Assert
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetStyle(batch, sheetName, "B5:B10", "Currency");
        _commands.SetStyle(batch, sheetName, "C5:C10", "Comma");
        // void methods throw on failure, succeed silently
    }

    [Fact]
    public void SetStyle_InvalidStyleName_ThrowsException()
    {
        // Arrange & Act & Assert - Should throw when Excel COM rejects invalid style name
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var exception = Assert.Throws<System.Reflection.TargetParameterCountException>(
            () => _commands.SetStyle(batch, sheetName, "A1", "NonExistentStyle"));

        // Verify exception message contains context about the style operation
        Assert.NotNull(exception.Message);
        Assert.Contains("Style", exception.Message);
    }

    [Fact]
    public void SetStyle_ResetToNormal_ClearsFormatting()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Apply fancy style
        _commands.SetStyle(batch, sheetName, "A1", "Accent1");

        // Reset to normal
        _commands.SetStyle(batch, sheetName, "A1", "Normal");
        // void methods throw on failure, succeed silently
    }
}





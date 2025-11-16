using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void SetStyle_Heading1_AppliesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_Heading1_AppliesSuccessfully),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.SetStyle(batch, "Sheet1", "A1", "Heading 1");

        // Assert
        Assert.True(result.Success, $"SetStyle failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void SetStyle_GoodBadNeutral_AllApplySuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_GoodBadNeutral_AllApplySuccessfully),
            _tempDir,
            ".xlsx");

        // Act & Assert
        using var batch = ExcelSession.BeginBatch(testFile);

        var goodResult = _commands.SetStyle(batch, "Sheet1", "A1", "Good");
        Assert.True(goodResult.Success, $"Good style failed: {goodResult.ErrorMessage}");

        var badResult = _commands.SetStyle(batch, "Sheet1", "A2", "Bad");
        Assert.True(badResult.Success, $"Bad style failed: {badResult.ErrorMessage}");

        var neutralResult = _commands.SetStyle(batch, "Sheet1", "A3", "Neutral");
        Assert.True(neutralResult.Success, $"Neutral style failed: {neutralResult.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void SetStyle_Accent1_AppliesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_Accent1_AppliesSuccessfully),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.SetStyle(batch, "Sheet1", "A1:E1", "Accent1");

        // Assert
        Assert.True(result.Success, $"Accent1 style failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void SetStyle_TotalStyle_AppliesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_TotalStyle_AppliesSuccessfully),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.SetStyle(batch, "Sheet1", "A10:E10", "Total");

        // Assert
        Assert.True(result.Success, $"Total style failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void SetStyle_CurrencyComma_AppliesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_CurrencyComma_AppliesSuccessfully),
            _tempDir,
            ".xlsx");

        // Act & Assert
        using var batch = ExcelSession.BeginBatch(testFile);

        var currencyResult = _commands.SetStyle(batch, "Sheet1", "B5:B10", "Currency");
        Assert.True(currencyResult.Success, $"Currency style failed: {currencyResult.ErrorMessage}");

        var commaResult = _commands.SetStyle(batch, "Sheet1", "C5:C10", "Comma");
        Assert.True(commaResult.Success, $"Comma style failed: {commaResult.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void SetStyle_InvalidStyleName_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_InvalidStyleName_ReturnsError),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.SetStyle(batch, "Sheet1", "A1", "NonExistentStyle");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("NonExistentStyle", result.ErrorMessage);
    }
    /// <inheritdoc/>

    [Fact]
    public void SetStyle_ResetToNormal_ClearsFormatting()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_ResetToNormal_ClearsFormatting),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Apply fancy style
        var fancyResult = _commands.SetStyle(batch, "Sheet1", "A1", "Accent1");
        Assert.True(fancyResult.Success);

        // Reset to normal
        var normalResult = _commands.SetStyle(batch, "Sheet1", "A1", "Normal");
        Assert.True(normalResult.Success, $"Normal style failed: {normalResult.ErrorMessage}");
    }
}

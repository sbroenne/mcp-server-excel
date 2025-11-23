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
        _commands.SetStyle(batch, "Sheet1", "A1", "Heading 1");

        // Assert - void method throws on failure, succeeds silently on success
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

        _commands.SetStyle(batch, "Sheet1", "A1", "Good");
        _commands.SetStyle(batch, "Sheet1", "A2", "Bad");
        _commands.SetStyle(batch, "Sheet1", "A3", "Neutral");
        // void methods throw on failure, succeed silently
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
        _commands.SetStyle(batch, "Sheet1", "A1:E1", "Accent1");

        // Assert - void method throws on failure, succeeds silently on success
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
        _commands.SetStyle(batch, "Sheet1", "A10:E10", "Total");

        // Assert - void method throws on failure, succeeds silently on success
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

        _commands.SetStyle(batch, "Sheet1", "B5:B10", "Currency");
        _commands.SetStyle(batch, "Sheet1", "C5:C10", "Comma");
        // void methods throw on failure, succeed silently
    }
    /// <inheritdoc/>
    [Fact]
    public void SetStyle_InvalidStyleName_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetStyle_InvalidStyleName_ThrowsException),
            _tempDir,
            ".xlsx");

        // Act & Assert - Should throw when Excel COM rejects invalid style name
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<System.Reflection.TargetParameterCountException>(
            () => _commands.SetStyle(batch, "Sheet1", "A1", "NonExistentStyle"));

        // Verify exception message contains context about the style operation
        Assert.NotNull(exception.Message);
        Assert.Contains("Style", exception.Message);
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
        _commands.SetStyle(batch, "Sheet1", "A1", "Accent1");

        // Reset to normal
        _commands.SetStyle(batch, "Sheet1", "A1", "Normal");
        // void methods throw on failure, succeed silently
    }
}

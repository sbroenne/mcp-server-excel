using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public async Task GetStyle_UnstyledRange_ReturnsNormalStyle()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(GetStyle_UnstyledRange_ReturnsNormalStyle),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.GetStyle(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"GetStyle failed: {result.ErrorMessage}");
        Assert.Equal("Normal", result.StyleName);
        Assert.True(result.IsBuiltInStyle);
        // Note: StyleDescription may be null for some styles
    }
    /// <inheritdoc/>

    [Fact]
    public async Task GetStyle_AfterSetStyle_ReturnsAppliedStyle()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(GetStyle_AfterSetStyle_ReturnsAppliedStyle),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Set a style first
        var setResult = _commands.SetStyle(batch, "Sheet1", "A1", "Heading 1");
        Assert.True(setResult.Success, $"SetStyle failed: {setResult.ErrorMessage}");

        // Now get the style
        var getResult = _commands.GetStyle(batch, "Sheet1", "A1");

        // Assert
        Assert.True(getResult.Success, $"GetStyle failed: {getResult.ErrorMessage}");
        Assert.Equal("Heading 1", getResult.StyleName);
        Assert.True(getResult.IsBuiltInStyle);
        // Note: StyleDescription may be null for some styles
    }
    /// <inheritdoc/>

    [Fact]
    public async Task GetStyle_MultipleStyles_ReturnsCorrectStyles()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(GetStyle_MultipleStyles_ReturnsCorrectStyles),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Set different styles on different cells
        var setHeading1 = await _commands.SetStyle(batch, "Sheet1", "A1", "Heading 1");
        var setAccent1 = await _commands.SetStyle(batch, "Sheet1", "B1", "Accent1");
        var setCurrency = await _commands.SetStyle(batch, "Sheet1", "C1", "Currency");

        Assert.True(setHeading1.Success, $"SetStyle Heading 1 failed: {setHeading1.ErrorMessage}");
        Assert.True(setAccent1.Success, $"SetStyle Accent1 failed: {setAccent1.ErrorMessage}");
        Assert.True(setCurrency.Success, $"SetStyle Currency failed: {setCurrency.ErrorMessage}");

        // Get the styles
        var getHeading1 = await _commands.GetStyle(batch, "Sheet1", "A1");
        var getAccent1 = await _commands.GetStyle(batch, "Sheet1", "B1");
        var getCurrency = await _commands.GetStyle(batch, "Sheet1", "C1");

        // Assert
        Assert.True(getHeading1.Success, $"GetStyle A1 failed: {getHeading1.ErrorMessage}");
        Assert.Equal("Heading 1", getHeading1.StyleName);
        Assert.True(getHeading1.IsBuiltInStyle);

        Assert.True(getAccent1.Success, $"GetStyle B1 failed: {getAccent1.ErrorMessage}");
        Assert.Equal("Accent1", getAccent1.StyleName);
        Assert.True(getAccent1.IsBuiltInStyle);

        Assert.True(getCurrency.Success, $"GetStyle C1 failed: {getCurrency.ErrorMessage}");
        Assert.Equal("Currency", getCurrency.StyleName);
        Assert.True(getCurrency.IsBuiltInStyle);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task GetStyle_RangeMultipleCells_ReturnsFirstCellStyle()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(GetStyle_RangeMultipleCells_ReturnsFirstCellStyle),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Set style on entire range (this applies to all cells in the range)
        var setResult = _commands.SetStyle(batch, "Sheet1", "A1:C3", "Good");
        Assert.True(setResult.Success, $"SetStyle failed: {setResult.ErrorMessage}");

        // Get style for entire range (should return first cell's style)
        var getResult = _commands.GetStyle(batch, "Sheet1", "A1:C3");

        // Assert
        Assert.True(getResult.Success, $"GetStyle failed: {getResult.ErrorMessage}");
        Assert.Equal("Good", getResult.StyleName);
        Assert.True(getResult.IsBuiltInStyle);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task GetStyle_InvalidRange_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(GetStyle_InvalidRange_ReturnsError),
            _tempDir,
            ".xlsx");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.GetStyle(batch, "Sheet1", "InvalidRange");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains(
            "range",
            result.ErrorMessage.ToLowerInvariant());
    }
}

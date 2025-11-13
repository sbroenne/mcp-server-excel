using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Tests for Named Range parameter name validation
/// Validates Excel's 255-character limit for named range names
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public partial class NamedRangeCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public async Task Create_EmptyParameterName_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Create_EmptyParameterName_ReturnsError), _tempDir);

        // Act
        var result = await _parameterCommands.CreateAsync(testFile, "", "Sheet1!A1");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_WhitespaceParameterName_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Create_WhitespaceParameterName_ReturnsError), _tempDir);

        // Act
        var result = await _parameterCommands.CreateAsync(testFile, "   ", "Sheet1!A1");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_ParameterNameExactly255Characters_ReturnsSuccess()
    {
        // Arrange - Create name with exactly 255 characters (Excel's limit)
        var paramName = new string('A', 255);
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Create_ParameterNameExactly255Characters_ReturnsSuccess), _tempDir);

        // Act
        var result = await _parameterCommands.CreateAsync(testFile, paramName, "Sheet1!A1");

        // Assert
        Assert.True(result.Success, $"Expected success with 255-char name but got error: {result.ErrorMessage}");

        // Verify the parameter was actually created
        var listResult = await _parameterCommands.ListAsync(testFile);
        Assert.Contains(listResult.NamedRanges, p => p.Name == paramName);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_ParameterName256Characters_ReturnsError()
    {
        // Arrange - Create name with 256 characters (exceeds Excel's limit)
        var paramName = new string('B', 256);
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Create_ParameterName256Characters_ReturnsError), _tempDir);

        // Act
        var result = await _parameterCommands.CreateAsync(testFile, paramName, "Sheet1!A1");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("255-character limit", result.ErrorMessage);
        Assert.Contains("256", result.ErrorMessage); // Should show actual length
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Update_ParameterNameExceeds255Characters_ReturnsError()
    {
        // Arrange
        var longParamName = new string('C', 300);
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Update_ParameterNameExceeds255Characters_ReturnsError), _tempDir);

        // Act
        var result = await _parameterCommands.UpdateAsync(testFile, longParamName, "Sheet1!B2");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("255-character limit", result.ErrorMessage);
        Assert.Contains("300", result.ErrorMessage);
    }
}

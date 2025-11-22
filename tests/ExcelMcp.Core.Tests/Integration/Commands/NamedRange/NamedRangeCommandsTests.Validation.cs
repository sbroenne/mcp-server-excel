using Sbroenne.ExcelMcp.ComInterop.Session;
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
    public void Create_EmptyParameterName_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Create_EmptyParameterName_ReturnsError), _tempDir);

        // Act & Assert - Empty parameter name should throw ArgumentException
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Create(batch, "", "Sheet1!A1"));

        Assert.Contains("cannot be empty", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_WhitespaceParameterName_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Create_WhitespaceParameterName_ReturnsError), _tempDir);

        // Act & Assert - Whitespace parameter name should throw ArgumentException
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Create(batch, "   ", "Sheet1!A1"));

        Assert.Contains("cannot be empty", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_ParameterNameExactly255Characters_ReturnsSuccess()
    {
        // Arrange - Create name with exactly 255 characters (Excel's limit)
        var paramName = new string('A', 255);
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Create_ParameterNameExactly255Characters_ReturnsSuccess), _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _parameterCommands.Create(batch, paramName, "Sheet1!A1");

        // Assert
        Assert.True(result.Success, $"Expected success with 255-char name but got error: {result.ErrorMessage}");

        // Verify the parameter was actually created
        var listResult = _parameterCommands.List(batch);
        Assert.Contains(listResult.NamedRanges, p => p.Name == paramName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_ParameterName256Characters_ReturnsError()
    {
        // Arrange - Create name with 256 characters (exceeds Excel's limit)
        var paramName = new string('B', 256);
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Create_ParameterName256Characters_ReturnsError), _tempDir);

        // Act & Assert - 256-character name should throw ArgumentException
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Create(batch, paramName, "Sheet1!A1"));

        Assert.Contains("255-character limit", exception.Message);
        Assert.Contains("256", exception.Message); // Should show actual length
    }
    /// <inheritdoc/>

    [Fact]
    public void Update_ParameterNameExceeds255Characters_ReturnsError()
    {
        // Arrange
        var longParamName = new string('C', 300);
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Update_ParameterNameExceeds255Characters_ReturnsError), _tempDir);

        // Act & Assert - 300-character name should throw ArgumentException
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Update(batch, longParamName, "Sheet1!B2"));

        Assert.Contains("255-character limit", exception.Message);
        Assert.Contains("300", exception.Message);
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateBulk_ParameterNameExceeds255Characters_SkipsWithError()
    {
        // Arrange
        var longParamName = new string('D', 270);
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(CreateBulk_ParameterNameExceeds255Characters_SkipsWithError), _tempDir);

        var parameters = new[]
        {
            new Models.NamedRangeDefinition { Name = "ValidParam1", Reference = "Sheet1!A1" },
            new Models.NamedRangeDefinition { Name = longParamName, Reference = "Sheet1!A2" },
            new Models.NamedRangeDefinition { Name = "ValidParam2", Reference = "Sheet1!A3" }
        };

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _parameterCommands.CreateBulk(batch, parameters);

        // Assert - Should succeed for valid params but skip the long one
        Assert.True(result.Success, $"Expected partial success but got error: {result.ErrorMessage}");

        // Verify valid parameters were created
        var listResult = _parameterCommands.List(batch);
        Assert.Contains(listResult.NamedRanges, p => p.Name == "ValidParam1");
        Assert.Contains(listResult.NamedRanges, p => p.Name == "ValidParam2");
        Assert.DoesNotContain(listResult.NamedRanges, p => p.Name == longParamName);
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateBulk_AllParametersExceedLimit_ReturnsError()
    {
        // Arrange
        var longParamName1 = new string('E', 260);
        var longParamName2 = new string('F', 280);
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(CreateBulk_AllParametersExceedLimit_ReturnsError), _tempDir);

        var parameters = new[]
        {
            new Models.NamedRangeDefinition { Name = longParamName1, Reference = "Sheet1!A1" },
            new Models.NamedRangeDefinition { Name = longParamName2, Reference = "Sheet1!A2" }
        };

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _parameterCommands.CreateBulk(batch, parameters);

        // Assert - Should fail because all parameters are invalid
        Assert.False(result.Success);
        Assert.Contains("Failed to create any parameters", result.ErrorMessage);
    }
}

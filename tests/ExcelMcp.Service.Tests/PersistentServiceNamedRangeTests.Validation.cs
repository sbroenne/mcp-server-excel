using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for Named Range parameter name validation
/// Validates Excel's 255-character limit for named range names
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceNamedRangeTests
{
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_EmptyParameterName_ThrowsHelpfulPublicError(string name)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Create(
                _fixture.BatchToken,
                name,
                "Sheet1!A1"));

        Assert.Contains(
            "name is required",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void Create_ParameterNameExactly255Characters_ReturnsSuccess()
    {
        // Arrange - Create name with exactly 255 characters (Excel's limit)
        // Named ranges must start with letter or underscore, so use "NR_" prefix
        var uniquePrefix = "NR_" + Guid.NewGuid().ToString("N")[..5] + "_";
        var paramName = uniquePrefix + new string('A', 255 - uniquePrefix.Length);

        // Act
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _parameterCommands.Create(batch, paramName, $"'{sheetName}'!A1");
        _fixture.RegisterNamedRangeForCleanup(paramName);

        // Assert - Verify the parameter was actually created
        var namedRanges = _parameterCommands.List(batch);
        Assert.True(namedRanges.Success);
        Assert.Contains(namedRanges.NamedRanges, p => p.Name == paramName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_ParameterName256Characters_ReturnsError()
    {
        // Arrange - Create name with 256 characters (exceeds Excel's limit)
        var paramName = new string('B', 256);

        // Act & Assert - 256-character name should throw ArgumentException
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
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

        // Act & Assert - 300-character name should throw ArgumentException
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Update(batch, longParamName, "Sheet1!B2"));

        Assert.Contains("255-character limit", exception.Message);
        Assert.Contains("300", exception.Message);
    }
}


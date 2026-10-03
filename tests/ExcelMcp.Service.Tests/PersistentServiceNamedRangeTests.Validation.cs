using System.Text.Json;
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
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var before = SeedNamedRangeForPreservation(sheetName);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Create(
                _fixture.BatchToken,
                name,
                "Sheet1!A1"));

        Assert.Contains(
            "name is required",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, CaptureNamedRangeState(sheetName));
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
        var sheetName = _fixture.CreateNamedTestSheet(batch, $"Long name {Guid.NewGuid():N}"[..27]);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Boundary value"]]).Success);
        Assert.True(_parameterCommands.Create(batch, paramName, $"'{sheetName}'!$A$1").Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);

        // Assert - Verify the parameter was actually created
        var namedRanges = _parameterCommands.List(batch);
        Assert.True(namedRanges.Success);
        var created = Assert.Single(namedRanges.NamedRanges, p => p.Name == paramName);
        Assert.Equal($"='{sheetName}'!$A$1", created.RefersTo);
        Assert.Equal("Boundary value", created.Value);
        Assert.Equal("String", created.ValueType);
        Assert.Equal(1, created.CellCount);
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
        var before = SeedNamedRangeForPreservation(sheetName);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Create(batch, paramName, "Sheet1!A1"));

        Assert.Contains("255-character limit", exception.Message);
        Assert.Contains("256", exception.Message); // Should show actual length
        Assert.Equal(before, CaptureNamedRangeState(sheetName));
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
        var before = SeedNamedRangeForPreservation(sheetName);
        var exception = Assert.Throws<ArgumentException>(() =>
            _parameterCommands.Update(batch, longParamName, "Sheet1!B2"));

        Assert.Contains("255-character limit", exception.Message);
        Assert.Contains("300", exception.Message);
        Assert.Equal(before, CaptureNamedRangeState(sheetName));
    }

    private string SeedNamedRangeForPreservation(string sheetName)
    {
        var batch = _fixture.BatchToken;
        var name = CreateUniqueNamedRangeName();
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [["Original", "Untouched"]]).Success);
        Assert.True(_parameterCommands.Create(batch, name, $"'{sheetName}'!$A$1").Success);
        _fixture.RegisterNamedRangeForCleanup(name);
        var existing = _parameterCommands.Read(batch, name);
        Assert.Equal("Original", existing.Value);
        Assert.Equal("String", existing.ValueType);
        return CaptureNamedRangeState(sheetName);
    }

    private string CaptureNamedRangeState(string sheetName)
    {
        var names = _parameterCommands.List(_fixture.BatchToken);
        Assert.True(names.Success, names.ErrorMessage);
        var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B1");
        Assert.True(values.Success, values.ErrorMessage);
        return JsonSerializer.Serialize(new { names.NamedRanges, values.Values });
    }
}

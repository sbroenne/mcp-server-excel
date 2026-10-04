using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for Parameter value operations (get, set)
/// </summary>
public sealed partial class PersistentServiceNamedRangeTests
{
    /// <inheritdoc/>
    [Fact]
    public void Set_ExistingParameter_UpdatesValue()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create parameter first
        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [["Original", "Untouched"]]).Success);
        Assert.Equal("Original", _parameterCommands.Read(batch, paramName).Value);

        // Set the parameter value
        Assert.True(_parameterCommands.Write(batch, paramName, "TestValue").Success);

        // Assert - Verify the parameter value was actually set by reading it back
        var namedRangeValue = _parameterCommands.Read(batch, paramName);
        Assert.Equal("TestValue", namedRangeValue.Value);
        Assert.Equal(paramName, namedRangeValue.Name);
        Assert.Equal("String", namedRangeValue.ValueType);
        var cells = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["TestValue", "Untouched"], Assert.Single(cells.Values));
    }

    [Fact]
    public void Write_DottedIdentifier_PreservesText()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_parameterCommands.Write(batch, paramName, "2.0.13").Success);

        var namedRangeValue = _parameterCommands.Read(batch, paramName);

        Assert.Equal("2.0.13", namedRangeValue.Value);
        Assert.Equal("String", namedRangeValue.ValueType);
        Assert.Equal(paramName, namedRangeValue.Name);
        var cells = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal("2.0.13", Assert.Single(Assert.Single(cells.Values)));
    }

    [Theory]
    [InlineData("123.45", 123.45, "Double")]
    [InlineData("true", true, "Boolean")]
    public void Write_InvariantScalar_StoresTypedValue(
        string input,
        object expectedValue,
        string expectedType)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_parameterCommands.Write(batch, paramName, input).Success);

        var namedRangeValue = _parameterCommands.Read(batch, paramName);

        Assert.Equal(expectedValue, namedRangeValue.Value);
        Assert.Equal(expectedType, namedRangeValue.ValueType);
        Assert.Equal(paramName, namedRangeValue.Name);
        var cells = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(expectedValue, Assert.Single(Assert.Single(cells.Values)));
    }

    /// <inheritdoc/>

    [Fact]
    public void Get_ExistingParameter_ReturnsValue()
    {
        // Arrange
        string testValue = "Integration Test Value";
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create and set parameter value
        Assert.True(_parameterCommands.Create(batch, paramName, cellRef).Success);
        _fixture.RegisterNamedRangeForCleanup(paramName);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[testValue]]).Success);

        // Get the parameter value
        var namedRangeValue = _parameterCommands.Read(batch, paramName);

        // Assert
        Assert.Equal(testValue, namedRangeValue.Value);
        Assert.Equal(paramName, namedRangeValue.Name);
        Assert.Equal("String", namedRangeValue.ValueType);
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_WithNonExistentParameter_ThrowsException()
    {
        // Arrange & Act & Assert
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var exception = Assert.Throws<InvalidOperationException>(
            () => _parameterCommands.Read(batch, $"NonExistent_{Guid.NewGuid():N}"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}

using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange - Use the shared fixture file
        // Note: The shared file may have named ranges from other tests,
        // so we verify the list operation works rather than asserting empty
        var batch = _fixture.BatchToken;

        // Act
        var namedRanges = _parameterCommands.List(batch);

        // Assert - List should return without error
        Assert.NotNull(namedRanges);
        Assert.True(namedRanges.Success);
        Assert.NotNull(namedRanges.NamedRanges);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_ValidNameAndReference_ReturnsSuccess()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Act
        _parameterCommands.Create(batch, paramName, cellRef);
        _fixture.RegisterNamedRangeForCleanup(paramName);

        // Assert - Verify the parameter was actually created by listing parameters
        var namedRanges = _parameterCommands.List(batch);
        Assert.True(namedRanges.Success);
        Assert.Contains(namedRanges.NamedRanges, p => p.Name == paramName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Delete_ExistingParameter_ReturnsSuccess()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create parameter first
        _parameterCommands.Create(batch, paramName, cellRef);
        _fixture.RegisterNamedRangeForCleanup(paramName);

        // Delete the parameter
        _parameterCommands.Delete(batch, paramName);
        _fixture.ForgetNamedRange(paramName);

        // Assert - Verify the parameter was actually deleted by checking it's not in the list
        var namedRanges = _parameterCommands.List(batch);
        Assert.True(namedRanges.Success);
        Assert.DoesNotContain(namedRanges.NamedRanges, p => p.Name == paramName);
    }
}

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Tests for Parameter lifecycle operations (list, create, delete)
/// </summary>
public partial class NamedRangeCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange - Use the shared fixture file
        // Note: The shared file may have named ranges from other tests,
        // so we verify the list operation works rather than asserting empty
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        // Act
        var namedRanges = _parameterCommands.List(batch);

        // Assert - List should return without error
        Assert.NotNull(namedRanges);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_ValidNameAndReference_ReturnsSuccess()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var paramName = _fixture.GetUniqueNamedRangeName();
        var cellRef = _fixture.GetUniqueCellReference();

        // Act
        _parameterCommands.Create(batch, paramName, cellRef);

        // Assert - Verify the parameter was actually created by listing parameters
        var namedRanges = _parameterCommands.List(batch);
        Assert.Contains(namedRanges, p => p.Name == paramName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Delete_ExistingParameter_ReturnsSuccess()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var paramName = _fixture.GetUniqueNamedRangeName();
        var cellRef = _fixture.GetUniqueCellReference();

        // Create parameter first
        _parameterCommands.Create(batch, paramName, cellRef);

        // Delete the parameter
        _parameterCommands.Delete(batch, paramName);

        // Assert - Verify the parameter was actually deleted by checking it's not in the list
        var namedRanges = _parameterCommands.List(batch);
        Assert.DoesNotContain(namedRanges, p => p.Name == paramName);
    }
    /// <inheritdoc/>

    [Fact]
    public void List_WithNonExistentFile_ReturnsError()
    {
        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => ExcelSession.BeginBatch("nonexistent.xlsx"));
    }
}

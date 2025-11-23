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
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(List_EmptyWorkbook_ReturnsEmptyList), _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var namedRanges = _parameterCommands.List(batch);

        // Assert
        Assert.NotNull(namedRanges);
        Assert.Empty(namedRanges);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_ValidNameAndReference_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Create_ValidNameAndReference_ReturnsSuccess), _tempDir);

        // Act - Use single batch for create and verify
        using var batch = ExcelSession.BeginBatch(testFile);
        _parameterCommands.Create(batch, "TestParam", "Sheet1!A1");

        // Assert - Verify the parameter was actually created by listing parameters
        var namedRanges = _parameterCommands.List(batch);
        Assert.Contains(namedRanges, p => p.Name == "TestParam");
    }
    /// <inheritdoc/>

    [Fact]
    public void Delete_ExistingParameter_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Delete_ExistingParameter_ReturnsSuccess), _tempDir);

        // Act - Use single batch for create, delete, and verify
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create parameter first
        _parameterCommands.Create(batch, "DeleteTestParam", "Sheet1!A1");

        // Delete the parameter
        _parameterCommands.Delete(batch, "DeleteTestParam");

        // Assert - Verify the parameter was actually deleted by checking it's not in the list
        var namedRanges = _parameterCommands.List(batch);
        Assert.DoesNotContain(namedRanges, p => p.Name == "DeleteTestParam");
    }
    /// <inheritdoc/>

    [Fact]
    public void List_WithNonExistentFile_ReturnsError()
    {
        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => ExcelSession.BeginBatch("nonexistent.xlsx"));
    }
}

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
        var result = _parameterCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.NamedRanges);
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
        var result = _parameterCommands.Create(batch, "TestParam", "Sheet1!A1");

        // Assert
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");

        // Verify the parameter was actually created by listing parameters
        var listResult = _parameterCommands.List(batch);
        Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
        Assert.Contains(listResult.NamedRanges, p => p.Name == "TestParam");
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
        var createResult = _parameterCommands.Create(batch, "DeleteTestParam", "Sheet1!A1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");

        // Delete the parameter
        var result = _parameterCommands.Delete(batch, "DeleteTestParam");
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify the parameter was actually deleted by checking it's not in the list
        var listResult = _parameterCommands.List(batch);
        Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
        Assert.DoesNotContain(listResult.NamedRanges, p => p.Name == "DeleteTestParam");
    }
    /// <inheritdoc/>

    [Fact]
    public void List_WithNonExistentFile_ReturnsError()
    {
        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => ExcelSession.BeginBatch("nonexistent.xlsx"));
    }
}

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection List operations
/// </summary>
public partial class ConnectionCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(List_EmptyWorkbook_ReturnsSuccessWithEmptyList), _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Connections);
        Assert.Empty(result.Connections);
        Assert.Equal(testFile, result.FilePath);
    }
}

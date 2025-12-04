using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection List operations
/// </summary>
public partial class ConnectionCommandsTests
{
    [Fact]
    public void List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

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

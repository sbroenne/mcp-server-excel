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
    public async Task List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
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
    /// <inheritdoc/>

    [Fact]
    public async Task List_WithTextConnection_ReturnsConnection()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ConnectionCommandsTests), nameof(List_WithTextConnection_ReturnsConnection), _tempDir);
        var csvFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ConnectionCommandsTests), nameof(List_WithTextConnection_ReturnsConnection), _tempDir, ".csv", "Name,Value\nTest1,100\nTest2,200");
        string connName = "TestText";

        await ConnectionTestHelper.CreateTextFileConnectionAsync(testFile, connName, csvFile);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.Connections);
        var conn = Assert.Single(result.Connections);
        Assert.Equal(connName, conn.Name);
        // Excel reports CSV files as WEB (type 4) instead of TEXT (type 3) - this is Excel's behavior
        Assert.Equal("WEB", conn.Type);
    }
}

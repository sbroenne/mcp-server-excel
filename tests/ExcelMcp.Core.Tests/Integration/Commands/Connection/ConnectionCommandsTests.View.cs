using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection View/Properties operations
/// </summary>
public partial class ConnectionCommandsTests
{
    [Fact]
    public async Task View_ExistingConnection_ReturnsDetails()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ConnectionCommandsTests), nameof(View_ExistingConnection_ReturnsDetails), _tempDir);
        var csvFile = CreateTestCsvFile($"View_{Guid.NewGuid():N}.csv");
        string connName = "ViewTestConnection";

        await ConnectionTestHelper.CreateTextFileConnectionAsync(testFile, connName, csvFile);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.ViewAsync(batch, connName);

        // Assert
        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(connName, result.ConnectionName);
        Assert.NotNull(result.ConnectionString);
        Assert.NotNull(result.Type);
    }

    [Fact]
    public async Task View_NonExistentConnection_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ConnectionCommandsTests), nameof(View_NonExistentConnection_ReturnsError), _tempDir);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _commands.ViewAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}

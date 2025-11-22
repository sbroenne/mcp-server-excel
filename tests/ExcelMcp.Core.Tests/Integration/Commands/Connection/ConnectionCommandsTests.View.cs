using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using System.IO;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection View/Properties operations
/// </summary>
public partial class ConnectionCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void View_ExistingConnection_ReturnsDetails()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(View_ExistingConnection_ReturnsDetails), _tempDir);

        // Use ODBC connection (doesn't need actual DSN for view test)
        var connName = "ViewTestConnection";
        string connectionString = "ODBC;DSN=ViewTestDSN;DBQ=C:\\temp\\viewtest.xlsx";
        ConnectionTestHelper.CreateOdbcConnection(testFile, connName, connectionString);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.View(batch, connName);

        // Assert
        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(connName, result.ConnectionName);
        Assert.NotNull(result.ConnectionString);
        Assert.NotNull(result.Type);
    }
    /// <inheritdoc/>

    [Fact]
    public void View_NonExistentConnection_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(View_NonExistentConnection_ReturnsError), _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.View(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}

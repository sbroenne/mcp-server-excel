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
    public void View_ExistingConnection_ReturnsDetails()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

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

    [Fact]
    public void View_NonExistentConnection_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act & Assert
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<InvalidOperationException>(() => _commands.View(batch, "NonExistent"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}

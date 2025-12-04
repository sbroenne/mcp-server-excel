using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection Delete operations
/// </summary>
public partial class ConnectionCommandsTests
{
    [Fact]
    public void Delete_ExistingTextConnection_ReturnsSuccess()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Use ODBC connection (doesn't need actual DSN for delete test)
        string connectionString = "ODBC;DSN=TestDSN;DBQ=C:\\temp\\test.xlsx";
        string connectionName = "DeleteTestConnection";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection first
        _commands.Create(batch, connectionName, connectionString);

        // Verify connection exists
        var listResultBefore = _commands.List(batch);
        Assert.True(listResultBefore.Success);
        Assert.Contains(listResultBefore.Connections, c => c.Name == connectionName);

        // Act - Delete the connection
        // Assert
        _commands.Delete(batch, connectionName);

        // Verify connection no longer exists
        var listResultAfter = _commands.List(batch);
        Assert.True(listResultAfter.Success);
        Assert.DoesNotContain(listResultAfter.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_NonExistentConnection_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "NonExistentConnection";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - Attempting to delete non-existent connection should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, connectionName);
        });

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Delete_AfterCreatingMultiple_RemovesOnlySpecified()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Use ODBC connections (don't need actual DSNs for delete test)
        string conn1Name = "Connection1";
        string conn2Name = "Connection2";
        string conn3Name = "Connection3";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create three connections
        _commands.Create(batch, conn1Name, "ODBC;DSN=TestDSN1;DBQ=C:\\temp\\test1.xlsx");
        _commands.Create(batch, conn2Name, "ODBC;DSN=TestDSN2;DBQ=C:\\temp\\test2.xlsx");
        _commands.Create(batch, conn3Name, "ODBC;DSN=TestDSN3;DBQ=C:\\temp\\test3.xlsx");

        // Act - Delete only the second connection
        // Assert
        _commands.Delete(batch, conn2Name);

        // Verify only conn2 is deleted
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == conn1Name);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == conn2Name);
        Assert.Contains(listResult.Connections, c => c.Name == conn3Name);
    }

    [Fact]
    public void Delete_ConnectionWithDescription_RemovesSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "DescribedConnection";
        string description = "Test connection with description";
        string connectionString = "ODBC;DSN=DescribedDSN;DBQ=C:\\temp\\described.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection with description
        _commands.Create(batch, connectionName, connectionString, null, description);

        // Act - Delete connection
        // Assert
        _commands.Delete(batch, connectionName);

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_ImmediatelyAfterCreate_WorksCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "ImmediateDeleteTest";
        string connectionString = "ODBC;DSN=ImmediateDSN;DBQ=C:\\temp\\immediate.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create and immediately delete
        _commands.Create(batch, connectionName, connectionString);

        // Assert
        _commands.Delete(batch, connectionName);

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_ConnectionAfterViewOperation_RemovesSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "ViewThenDelete";
        string connectionString = "ODBC;DSN=ViewDeleteDSN;DBQ=C:\\temp\\viewdelete.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create and view connection
        _commands.Create(batch, connectionName, connectionString);

        var viewResult = _commands.View(batch, connectionName);
        Assert.True(viewResult.Success);
        Assert.Equal(connectionName, viewResult.ConnectionName);

        // Act - Delete after viewing
        // Assert
        _commands.Delete(batch, connectionName);

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_EmptyConnectionName_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - Empty connection name should throw
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, string.Empty);
        });

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Delete_RepeatedDeleteAttempts_SecondAttemptFails()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionName = "DoubleDeleteTest";
        string connectionString = "ODBC;DSN=DoubleDeleteDSN;DBQ=C:\\temp\\doubledelete.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection
        _commands.Create(batch, connectionName, connectionString);

        // Act - First delete
        _commands.Delete(batch, connectionName);

        // Act & Assert - Second delete should fail
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, connectionName);
        });

        Assert.Contains("not found", exception.Message);
    }
}

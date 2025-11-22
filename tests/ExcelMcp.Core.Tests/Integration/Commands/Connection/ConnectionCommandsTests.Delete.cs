using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
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
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_ExistingTextConnection_ReturnsSuccess),
            _tempDir);

        // Use ODBC connection (doesn't need actual DSN for delete test)
        string connectionString = "ODBC;DSN=TestDSN;DBQ=C:\\temp\\test.xlsx";
        string connectionName = "DeleteTestConnection";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection first
        var createResult = _commands.Create(batch, connectionName, connectionString);
        Assert.True(createResult.Success, $"Connection creation failed: {createResult.ErrorMessage}");

        // Verify connection exists
        var listResultBefore = _commands.List(batch);
        Assert.True(listResultBefore.Success);
        Assert.Contains(listResultBefore.Connections, c => c.Name == connectionName);

        // Act - Delete the connection
        var deleteResult = _commands.Delete(batch, connectionName);

        // Assert
        Assert.True(deleteResult.Success, $"Delete operation failed: {deleteResult.ErrorMessage}");
        Assert.Null(deleteResult.ErrorMessage);

        // Verify connection no longer exists
        var listResultAfter = _commands.List(batch);
        Assert.True(listResultAfter.Success);
        Assert.DoesNotContain(listResultAfter.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_NonExistentConnection_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_NonExistentConnection_ThrowsException),
            _tempDir);

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
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_AfterCreatingMultiple_RemovesOnlySpecified),
            _tempDir);

        // Use ODBC connections (don't need actual DSNs for delete test)
        string conn1Name = "Connection1";
        string conn2Name = "Connection2";
        string conn3Name = "Connection3";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create three connections
        var result1 = _commands.Create(batch, conn1Name, "ODBC;DSN=TestDSN1;DBQ=C:\\temp\\test1.xlsx");
        var result2 = _commands.Create(batch, conn2Name, "ODBC;DSN=TestDSN2;DBQ=C:\\temp\\test2.xlsx");
        var result3 = _commands.Create(batch, conn3Name, "ODBC;DSN=TestDSN3;DBQ=C:\\temp\\test3.xlsx");

        Assert.True(result1.Success);
        Assert.True(result2.Success);
        Assert.True(result3.Success);

        // Act - Delete only the second connection
        var deleteResult = _commands.Delete(batch, conn2Name);

        // Assert
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

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
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_ConnectionWithDescription_RemovesSuccessfully),
            _tempDir);

        string connectionName = "DescribedConnection";
        string description = "Test connection with description";
        string connectionString = "ODBC;DSN=DescribedDSN;DBQ=C:\\temp\\described.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection with description
        var createResult = _commands.Create(batch, connectionName, connectionString, null, description);
        Assert.True(createResult.Success);

        // Act - Delete connection
        var deleteResult = _commands.Delete(batch, connectionName);

        // Assert
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_ImmediatelyAfterCreate_WorksCorrectly()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_ImmediatelyAfterCreate_WorksCorrectly),
            _tempDir);

        string connectionName = "ImmediateDeleteTest";
        string connectionString = "ODBC;DSN=ImmediateDSN;DBQ=C:\\temp\\immediate.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create and immediately delete
        var createResult = _commands.Create(batch, connectionName, connectionString);
        Assert.True(createResult.Success);

        var deleteResult = _commands.Delete(batch, connectionName);

        // Assert
        Assert.True(deleteResult.Success, $"Immediate delete failed: {deleteResult.ErrorMessage}");

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_ConnectionAfterViewOperation_RemovesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_ConnectionAfterViewOperation_RemovesSuccessfully),
            _tempDir);

        string connectionName = "ViewThenDelete";
        string connectionString = "ODBC;DSN=ViewDeleteDSN;DBQ=C:\\temp\\viewdelete.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create and view connection
        var createResult = _commands.Create(batch, connectionName, connectionString);
        Assert.True(createResult.Success);

        var viewResult = _commands.View(batch, connectionName);
        Assert.True(viewResult.Success);
        Assert.Equal(connectionName, viewResult.ConnectionName);

        // Act - Delete after viewing
        var deleteResult = _commands.Delete(batch, connectionName);

        // Assert
        Assert.True(deleteResult.Success, $"Delete after view failed: {deleteResult.ErrorMessage}");

        var listResult = _commands.List(batch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Delete_EmptyConnectionName_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_EmptyConnectionName_ThrowsException),
            _tempDir);

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
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Delete_RepeatedDeleteAttempts_SecondAttemptFails),
            _tempDir);

        string connectionName = "DoubleDeleteTest";
        string connectionString = "ODBC;DSN=DoubleDeleteDSN;DBQ=C:\\temp\\doubledelete.xlsx";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection
        var createResult = _commands.Create(batch, connectionName, connectionString);
        Assert.True(createResult.Success);

        // Act - First delete
        var firstDeleteResult = _commands.Delete(batch, connectionName);
        Assert.True(firstDeleteResult.Success);

        // Act & Assert - Second delete should fail
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            _commands.Delete(batch, connectionName);
        });

        Assert.Contains("not found", exception.Message);
    }
}

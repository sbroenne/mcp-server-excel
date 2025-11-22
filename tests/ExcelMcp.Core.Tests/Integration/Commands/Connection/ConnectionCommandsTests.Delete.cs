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

        // Create a CSV file to connect to
        var csvPath = Path.Combine(_tempDir, "delete_test_data.csv");
        System.IO.File.WriteAllText(csvPath, "Name,Value\nTest,123");

        string connectionString = $"TEXT;{csvPath}";
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

        // Create multiple CSV files
        var csv1Path = Path.Combine(_tempDir, "multi_delete_1.csv");
        var csv2Path = Path.Combine(_tempDir, "multi_delete_2.csv");
        var csv3Path = Path.Combine(_tempDir, "multi_delete_3.csv");
        System.IO.File.WriteAllText(csv1Path, "A,B\n1,2");
        System.IO.File.WriteAllText(csv2Path, "C,D\n3,4");
        System.IO.File.WriteAllText(csv3Path, "E,F\n5,6");

        string conn1Name = "Connection1";
        string conn2Name = "Connection2";
        string conn3Name = "Connection3";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create three connections
        var result1 = _commands.Create(batch, conn1Name, $"TEXT;{csv1Path}");
        var result2 = _commands.Create(batch, conn2Name, $"TEXT;{csv2Path}");
        var result3 = _commands.Create(batch, conn3Name, $"TEXT;{csv3Path}");

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

        var csvPath = Path.Combine(_tempDir, "described_connection.csv");
        System.IO.File.WriteAllText(csvPath, "Name,Value\nTest,100");

        string connectionName = "DescribedConnection";
        string description = "Test connection with description";
        string connectionString = $"TEXT;{csvPath}";

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

        var csvPath = Path.Combine(_tempDir, "immediate_delete.csv");
        System.IO.File.WriteAllText(csvPath, "X,Y\n10,20");

        string connectionName = "ImmediateDeleteTest";
        string connectionString = $"TEXT;{csvPath}";

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

        var csvPath = Path.Combine(_tempDir, "view_then_delete.csv");
        System.IO.File.WriteAllText(csvPath, "Col1,Col2\nVal1,Val2");

        string connectionName = "ViewThenDelete";
        string connectionString = $"TEXT;{csvPath}";

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

        var csvPath = Path.Combine(_tempDir, "double_delete.csv");
        System.IO.File.WriteAllText(csvPath, "A,B\n1,2");

        string connectionName = "DoubleDeleteTest";
        string connectionString = $"TEXT;{csvPath}";

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

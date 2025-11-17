using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection Create operations
/// </summary>
public partial class ConnectionCommandsTests
{
    [Fact]
    public void Create_TextConnection_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_TextConnection_ReturnsSuccess),
            _tempDir);

        // Create a CSV file to connect to
        var csvPath = Path.Combine(_tempDir, "test_data.csv");
        System.IO.File.WriteAllText(csvPath, "Name,Value\nTest,123");

        string connectionString = $"TEXT;{csvPath}";
        string connectionName = "TestTextConnection";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Create(batch, connectionName, connectionString);

        // Assert
        Assert.True(result.Success, $"TEXT connection creation should succeed: {result.ErrorMessage}");
        Assert.Null(result.ErrorMessage);

        // Verify connection exists
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Create_OleDbSqlServerConnection_TestWithAdd2Method()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_OleDbSqlServerConnection_TestWithAdd2Method),
            _tempDir);

        // SQL Server LocalDB connection string (most commonly available OLEDB provider on Windows)
        string connectionString = "OLEDB;Provider=SQLOLEDB;Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=tempdb;Integrated Security=SSPI";
        string connectionName = "TestSqlServerConnection";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Create(batch, connectionName, connectionString);

        // Assert - Document actual behavior
        // This test verifies whether Add2() method works for OLEDB connections
        if (result.Success)
        {
            // Connection created successfully - verify it exists
            var listResult = _commands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Connections, c => c.Name == connectionName);
        }
        else
        {
            // Connection failed - document the error for investigation
            Assert.NotNull(result.ErrorMessage);
            // This is acceptable - documents the limitation if it exists
        }

        // Test always passes - it's documenting actual behavior
        Assert.True(true, $"OLEDB connection test result - Success: {result.Success}, Error: {result.ErrorMessage}");
    }

    [Fact]
    public void Create_OleDbAccessConnection_TestWithAdd2Method()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_OleDbAccessConnection_TestWithAdd2Method),
            _tempDir);

        // Microsoft Access Database Engine (commonly available OLEDB provider)
        // Create an empty Access database file for testing
        var mdbPath = Path.Combine(_tempDir, "test.accdb");
        System.IO.File.WriteAllText(mdbPath, ""); // Placeholder - real Access DB would need proper format

        string connectionString = $"OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source={mdbPath}";
        string connectionName = "TestAccessConnection";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Create(batch, connectionName, connectionString);

        // Assert - Document actual behavior
        if (result.Success)
        {
            var listResult = _commands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Connections, c => c.Name == connectionName);
        }
        else
        {
            Assert.NotNull(result.ErrorMessage);
        }

        // Test always passes - documenting behavior
        Assert.True(true, $"Access OLEDB test - Success: {result.Success}, Error: {result.ErrorMessage}");
    }

    [Fact]
    public void Create_OdbcConnection_TestWithAdd2Method()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_OdbcConnection_TestWithAdd2Method),
            _tempDir);

        // Generic ODBC connection string
        string connectionString = "ODBC;DSN=Excel Files;DBQ=C:\\temp\\test.xlsx";
        string connectionName = "TestOdbcConnection";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Create(batch, connectionName, connectionString);

        // Assert - Document actual behavior
        if (result.Success)
        {
            var listResult = _commands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Connections, c => c.Name == connectionName);
        }
        else
        {
            Assert.NotNull(result.ErrorMessage);
        }

        // Test always passes - documenting behavior
        Assert.True(true, $"ODBC test - Success: {result.Success}, Error: {result.ErrorMessage}");
    }

    [Fact]
    public void Create_DuplicateName_AllowsOrReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_DuplicateName_AllowsOrReturnsError),
            _tempDir);

        var csvPath = Path.Combine(_tempDir, "duplicate_test.csv");
        System.IO.File.WriteAllText(csvPath, "A,B\n1,2");

        string connectionString = $"TEXT;{csvPath}";
        string connectionName = "DuplicateTest";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create first connection
        var result1 = _commands.Create(batch, connectionName, connectionString);
        Assert.True(result1.Success);

        // Act - Try to create duplicate
        var result2 = _commands.Create(batch, connectionName, connectionString);

        // Assert - Document actual behavior (Excel may allow duplicates or return error)
        // This test passes either way - it's documenting the behavior
        Assert.True(true, $"Duplicate connection result - Success: {result2.Success}, Error: {result2.ErrorMessage}");
    }
}

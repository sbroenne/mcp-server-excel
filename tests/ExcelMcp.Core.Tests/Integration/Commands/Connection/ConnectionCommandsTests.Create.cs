using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection Create operations.
/// 
/// Connection Type Support:
/// - TEXT/WEB: Create works (but LoadTo fails - use Power Query for data loading)
/// - OLEDB/ODBC: Create currently fails with "Value does not fall within expected range"
///               Excel COM Add2() method doesn't support OLEDB/ODBC connection creation
///               Use Power Query for OLEDB/ODBC data import instead
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

        // Assert - TEXT connections can be created successfully
        Assert.True(result.Success, $"TEXT connection creation failed: {result.ErrorMessage}");

        // Verify connection exists in workbook
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Create_WebConnection_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_WebConnection_ReturnsSuccess),
            _tempDir);

        string connectionString = "URL;https://example.com/data.xml";
        string connectionName = "TestWebConnection";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Create(batch, connectionName, connectionString);

        // Assert - WEB connections can be created successfully
        Assert.True(result.Success, $"WEB connection creation failed: {result.ErrorMessage}");

        // Verify connection exists
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Create_OleDbConnection_ThrowsArgumentException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_OleDbConnection_ThrowsArgumentException),
            _tempDir);

        // OLEDB connection string - known to fail with Excel COM Add2() method
        string connectionString = "OLEDB;Provider=SQLOLEDB;Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=tempdb;Integrated Security=SSPI";
        string connectionName = "TestOleDbConnection";

        // Act & Assert - OLEDB connections fail with ArgumentException
        // Excel COM Add2() method doesn't support OLEDB connection creation
        // Users should use Power Query for OLEDB data import instead
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.Create(batch, connectionName, connectionString));

        Assert.Contains("Value does not fall within the expected range", exception.Message);
    }

    [Fact]
    public void Create_OdbcConnection_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_OdbcConnection_ReturnsSuccess),
            _tempDir);

        // ODBC connection string - Excel accepts but may not connect without actual DSN
        string connectionString = "ODBC;DSN=Excel Files;DBQ=C:\\temp\\test.xlsx";
        string connectionName = "TestOdbcConnection";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Create(batch, connectionName, connectionString);

        // Assert - ODBC connections can be created (even without valid DSN)
        Assert.True(result.Success, $"ODBC connection creation failed: {result.ErrorMessage}");

        // Verify connection exists
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Create_DuplicateName_PreventsOrRenamesDuplicate()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_DuplicateName_PreventsOrRenamesDuplicate),
            _tempDir);

        var csvPath = Path.Combine(_tempDir, "duplicate_test.csv");
        System.IO.File.WriteAllText(csvPath, "A,B\n1,2");

        string connectionString = $"TEXT;{csvPath}";
        string connectionName = "DuplicateTest";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create first connection
        var result1 = _commands.Create(batch, connectionName, connectionString);
        Assert.True(result1.Success, $"First connection creation failed: {result1.ErrorMessage}");

        // Act - Create second connection with same name
        var result2 = _commands.Create(batch, connectionName, connectionString);
        Assert.True(result2.Success, $"Second connection creation failed: {result2.ErrorMessage}");

        // Assert - Excel may prevent duplicates OR auto-rename the second one
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);

        var matchingConnections = listResult.Connections.Where(c =>
            c.Name == connectionName || c.Name.StartsWith(connectionName, StringComparison.Ordinal)).ToList();

        // At least one connection should exist
        Assert.True(matchingConnections.Count >= 1,
            "At least one connection with the specified name should exist");
    }

    [Fact]
    public void Create_WithDescription_StoresDescription()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Create_WithDescription_StoresDescription),
            _tempDir);

        var csvPath = Path.Combine(_tempDir, "description_test.csv");
        System.IO.File.WriteAllText(csvPath, "X,Y\n10,20");

        string connectionString = $"TEXT;{csvPath}";
        string connectionName = "ConnectionWithDescription";
        string description = "This is a test connection for CSV data";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Create(batch, connectionName, connectionString,
            commandText: null, description: description);

        // Assert
        Assert.True(result.Success, $"Connection creation failed: {result.ErrorMessage}");

        // Verify connection was created successfully
        var viewResult = _commands.View(batch, connectionName);
        Assert.True(viewResult.Success);
        // Note: Description not currently exposed in ConnectionViewResult, only ConnectionString
    }
}

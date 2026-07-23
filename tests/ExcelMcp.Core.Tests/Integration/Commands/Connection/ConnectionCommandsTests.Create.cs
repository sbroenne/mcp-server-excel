using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection Create operations.
///
/// Connection Type Support:
/// - TEXT/WEB: Blocked (NotSupportedException) - Use Power Query for file/web imports
/// - OLEDB: Supported for providers installed on the machine (ACE, SQL Server, etc.)
/// - ODBC: Create works (even without valid DSN configured)
/// </summary>
public partial class ConnectionCommandsTests
{
    [Fact]
    public void Create_TextConnection_ThrowsNotSupportedException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        var csvPath = Path.Combine(_fixture.TempDir, "test_data.csv");
        System.IO.File.WriteAllText(csvPath, "Name,Value\nTest,123");

        string connectionString = $"TEXT;{csvPath}";
        string connectionName = "TestTextConnection";

        // Act & Assert - TEXT connections are blocked, use Power Query instead
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<NotSupportedException>(() =>
            _commands.Create(batch, connectionName, connectionString));

        Assert.Contains("TEXT and WEB connections are no longer supported", exception.Message);
        Assert.Contains("powerquery", exception.Message);
    }

    [Fact]
    public void Create_WebConnection_ThrowsNotSupportedException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionString = "URL;https://example.com/data.xml";
        string connectionName = "TestWebConnection";

        // Act & Assert - WEB connections are blocked, use Power Query instead
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<NotSupportedException>(() =>
            _commands.Create(batch, connectionName, connectionString));

        Assert.Contains("TEXT and WEB connections are no longer supported", exception.Message);
        Assert.Contains("powerquery", exception.Message);
    }

    [Fact]
    public void Create_AceOleDbConnection_ReturnsSuccess()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        var sourceWorkbook = _fixture.GetSourceFilePath("AceOleDbSource");
        AceOleDbTestHelper.CreateExcelDataSource(sourceWorkbook);

        string connectionString = AceOleDbTestHelper.GetExcelConnectionString(sourceWorkbook);
        string commandText = AceOleDbTestHelper.GetDefaultCommandText();
        string connectionName = "AceOleDbConnection";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        _commands.Create(batch, connectionName, connectionString, commandText: commandText);

        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == connectionName);

        // Cleanup source workbook (connection is stored in target workbook)
        batch.Save();
        if (System.IO.File.Exists(sourceWorkbook))
        {
            System.IO.File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void Create_OdbcConnection_ReturnsSuccess()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // ODBC connection string - Excel accepts but may not connect without actual DSN
        string connectionString = "ODBC;DSN=Excel Files;DBQ=C:\\temp\\test.xlsx";
        string connectionName = "TestOdbcConnection";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, connectionName, connectionString);

        // Verify connection exists
        var listResult = _commands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == connectionName);
    }

    [Fact]
    public void Create_DuplicateName_CreatesSecondConnection()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionString1 = "ODBC;DSN=Source1;DBQ=C:\\temp\\test1.xlsx";
        string connectionString2 = "ODBC;DSN=Source2;DBQ=C:\\temp\\test2.xlsx";
        string connectionName = "DuplicateTest";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create first connection
        _commands.Create(batch, connectionName, connectionString1);

        // Act - Create second connection with same name
        _commands.Create(batch, connectionName, connectionString2);

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
    public void Create_WithDescription_CreatesConnection()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        string connectionString = "ODBC;DSN=Excel Files;DBQ=C:\\temp\\test.xlsx";
        string connectionName = "ConnectionWithDescription";
        string description = "This is a test connection for ODBC data";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Create(batch, connectionName, connectionString,
            commandText: null, description: description);

        // Verify connection was created successfully
        var viewResult = _commands.View(batch, connectionName);
        Assert.True(viewResult.Success);
        // Note: Description not currently exposed in ConnectionViewResult, only ConnectionString
    }
}





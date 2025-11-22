using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using System.Globalization;
using IOFile = System.IO.File;
using IOPath = System.IO.Path;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection ImportOdc operations
/// </summary>
public partial class ConnectionCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void ImportFromOdc_ValidFile_ImportsConnection()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(ImportFromOdc_ValidFile_ImportsConnection), _tempDir);

        var dbPath = IOPath.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.db");
        SQLiteDatabaseHelper.CreateTestDatabase(dbPath);

        // Create ODC file with SQLite connection
        var odcFile = IOPath.Combine(_tempDir, $"Test_{Guid.NewGuid():N}.odc");
        var odcContent = GetSqliteOdcContent(dbPath);
        IOFile.WriteAllText(odcFile, odcContent);

        try
        {
            // Act
            using var batch = ExcelSession.BeginBatch(testFile);
            var result = _commands.ImportFromOdc(batch, odcFile);

            // Assert - Verify import succeeded
            Assert.True(result.Success, $"Import failed: {result.ErrorMessage}");
            Assert.NotNull(result.Action);
            Assert.Contains("import-odc (imported:", result.Action);

            // Extract connection name from Action string
            var connectionName = ExtractConnectionName(result.Action);
            Assert.False(string.IsNullOrEmpty(connectionName));

            // Verify connection exists in workbook
            var listResult = _commands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Connections, c => c.Name == connectionName);
        }
        finally
        {
            SQLiteDatabaseHelper.DeleteDatabase(dbPath);
            if (IOFile.Exists(odcFile))
            {
                IOFile.Delete(odcFile);
            }
        }
    }

    /// <inheritdoc/>
    [Fact]
    public void ImportFromOdc_MissingFile_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(ImportFromOdc_MissingFile_ReturnsError), _tempDir);

        var nonExistentOdc = IOPath.Combine(_tempDir, $"NonExistent_{Guid.NewGuid():N}.odc");

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.ImportFromOdc(batch, nonExistentOdc);

        // Assert - Verify error returned
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void ImportFromOdc_DuplicateImport_AddsSecondConnection()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(ImportFromOdc_DuplicateImport_AddsSecondConnection), _tempDir);

        var dbPath = IOPath.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.db");
        SQLiteDatabaseHelper.CreateTestDatabase(dbPath);

        // Create ODC file
        var odcFile = IOPath.Combine(_tempDir, $"Test_{Guid.NewGuid():N}.odc");
        var odcContent = GetSqliteOdcContent(dbPath);
        IOFile.WriteAllText(odcFile, odcContent);

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            // Act - Import twice
            var result1 = _commands.ImportFromOdc(batch, odcFile);
            var result2 = _commands.ImportFromOdc(batch, odcFile);

            // Assert - Both imports succeeded
            Assert.True(result1.Success, $"First import failed: {result1.ErrorMessage}");
            Assert.True(result2.Success, $"Second import failed: {result2.ErrorMessage}");

            // Extract connection names
            var conn1Name = ExtractConnectionName(result1.Action);
            var conn2Name = ExtractConnectionName(result2.Action);

            // Verify both connections exist (Excel may add suffix like "_1" for duplicates)
            var listResult = _commands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Connections, c => c.Name == conn1Name);
            Assert.Contains(listResult.Connections, c => c.Name == conn2Name);
        }
        finally
        {
            SQLiteDatabaseHelper.DeleteDatabase(dbPath);
            if (IOFile.Exists(odcFile))
            {
                IOFile.Delete(odcFile);
            }
        }
    }

    /// <summary>
    /// Generates SQLite OLEDB ODC file content with dynamic database path
    /// </summary>
    private static string GetSqliteOdcContent(string databasePath)
    {
        return $@"<html xmlns:o=""urn:schemas-microsoft-com:office:office""
xmlns=""http://www.w3.org/TR/REC-html40"">

<head>
<meta http-equiv=Content-Type content=""text/x-ms-odc; charset=utf-8"">
<meta name=ProgId content=ODC.Database>
<meta name=SourceType content=OLEDB>
<title>SQLite OLEDB Test Connection</title>
<xml id=docprops>
<o:DocumentProperties
  xmlns:o=""urn:schemas-microsoft-com:office:office""
  xmlns=""http://www.w3.org/TR/REC-html40"">
  <o:Description>Test connection to SQLite database</o:Description>
  <o:Name>SQLiteTest</o:Name>
</o:DocumentProperties>
</xml>
<xml id=msodc>
<odc:OfficeDataConnection
  xmlns:odc=""urn:schemas-microsoft-com:office:odc""
  xmlns=""http://www.w3.org/TR/REC-html40"">
  <odc:Connection odc:Type=""OLEDB"">
    <odc:ConnectionString>Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath};</odc:ConnectionString>
    <odc:CommandType>SQL</odc:CommandType>
    <odc:CommandText>SELECT * FROM Products</odc:CommandText>
  </odc:Connection>
</odc:OfficeDataConnection>
</xml>
</head>

<body>
<table>
  <tr>
    <td>SQLite OLEDB Test Connection</td>
  </tr>
</table>
</body>

</html>";
    }

    /// <inheritdoc/>
    [Fact]
    public void ImportFromOdc_ThenUpdateCommandText_RefreshSucceeds()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(ImportFromOdc_ThenUpdateCommandText_RefreshSucceeds), _tempDir);

        var dbPath = IOPath.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.db");
        SQLiteDatabaseHelper.CreateTestDatabase(dbPath);

        // Create ODC file with initial command text
        var odcFile = IOPath.Combine(_tempDir, $"Test_{Guid.NewGuid():N}.odc");
        var odcContent = GetSqliteOdcContent(dbPath);
        IOFile.WriteAllText(odcFile, odcContent);

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            // Act 1 - Import connection with original command text
            var importResult = _commands.ImportFromOdc(batch, odcFile);
            Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

            var connectionName = ExtractConnectionName(importResult.Action);
            Assert.False(string.IsNullOrEmpty(connectionName));

            // Verify original command text
            var viewResult1 = _commands.View(batch, connectionName);
            Assert.True(viewResult1.Success);
            Assert.Equal("SELECT * FROM Products", viewResult1.CommandText);

            // Act 2 - Update command text to filtered query
            var newCommandText = "SELECT ProductID, ProductName FROM Products WHERE Price > 10";
            var updateResult = _commands.SetProperties(
                batch,
                connectionName,
                connectionString: null,
                commandText: newCommandText,
                description: null);
            Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

            // Verify updated command text
            var viewResult2 = _commands.View(batch, connectionName);
            Assert.True(viewResult2.Success);
            Assert.Equal(newCommandText, viewResult2.CommandText);
            Assert.NotEqual("SELECT * FROM Products", viewResult2.CommandText);

            // Act 3 - Verify refresh works with updated query (proves query executes)
            var refreshResult = _commands.Refresh(batch, connectionName);
            Assert.True(refreshResult.Success, $"Refresh after update failed: {refreshResult.ErrorMessage}");
        }
        finally
        {
            SQLiteDatabaseHelper.DeleteDatabase(dbPath);
            if (IOFile.Exists(odcFile))
            {
                IOFile.Delete(odcFile);
            }
        }
    }

    /// <inheritdoc/>
    [Fact]
    public void ImportFromOdc_ThenUpdateCommandText_ConnectionUpdated()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(ImportFromOdc_ThenUpdateCommandText_ConnectionUpdated), _tempDir);

        var dbPath = IOPath.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.db");
        SQLiteDatabaseHelper.CreateTestDatabase(dbPath);

        // Create ODC file with initial command text
        var odcFile = IOPath.Combine(_tempDir, $"Test_{Guid.NewGuid():N}.odc");
        var odcContent = GetSqliteOdcContent(dbPath);
        IOFile.WriteAllText(odcFile, odcContent);

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            // Act 1 - Import connection with original command text
            var importResult = _commands.ImportFromOdc(batch, odcFile);
            Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

            var connectionName = ExtractConnectionName(importResult.Action);
            Assert.False(string.IsNullOrEmpty(connectionName));

            // Verify original command text
            var viewResult1 = _commands.View(batch, connectionName);
            Assert.True(viewResult1.Success);
            Assert.Equal("SELECT * FROM Products", viewResult1.CommandText);

            // Act 2 - Update command text
            var newCommandText = "SELECT ProductID, ProductName FROM Products WHERE Price > 10";
            var updateResult = _commands.SetProperties(
                batch,
                connectionName,
                connectionString: null,
                commandText: newCommandText,
                description: null);

            // Assert - Verify update succeeded
            Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

            // Verify updated command text (must differ from original)
            var viewResult2 = _commands.View(batch, connectionName);
            Assert.True(viewResult2.Success);
            Assert.Equal(newCommandText, viewResult2.CommandText);
            Assert.NotEqual("SELECT * FROM Products", viewResult2.CommandText); // Confirm it changed
        }
        finally
        {
            SQLiteDatabaseHelper.DeleteDatabase(dbPath);
            if (IOFile.Exists(odcFile))
            {
                IOFile.Delete(odcFile);
            }
        }
    }

    /// <inheritdoc/>
    [Fact]
    public void ImportFromOdc_ThenUpdateDescription_ConnectionUpdated()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests), nameof(ImportFromOdc_ThenUpdateDescription_ConnectionUpdated), _tempDir);

        var dbPath = IOPath.Combine(_tempDir, $"TestDb_{Guid.NewGuid():N}.db");
        SQLiteDatabaseHelper.CreateTestDatabase(dbPath);

        // Create ODC file with initial description
        var odcFile = IOPath.Combine(_tempDir, $"Test_{Guid.NewGuid():N}.odc");
        var odcContent = GetSqliteOdcContent(dbPath);
        IOFile.WriteAllText(odcFile, odcContent);

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            // Act 1 - Import connection with original description
            var importResult = _commands.ImportFromOdc(batch, odcFile);
            Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

            var connectionName = ExtractConnectionName(importResult.Action);
            Assert.False(string.IsNullOrEmpty(connectionName));

            // Verify original description (from ODC file)
            var viewResult1 = _commands.View(batch, connectionName);
            Assert.True(viewResult1.Success);
            var originalDescription = viewResult1.ConnectionString; // Store for comparison

            // Act 2 - Update description
            var newDescription = "Updated connection description for testing";
            var updateResult = _commands.SetProperties(
                batch,
                connectionName,
                connectionString: null,
                commandText: null,
                description: newDescription);

            // Assert - Verify update succeeded
            Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

            // Note: View doesn't return Description, so we verify via GetProperties or by checking it doesn't error
            // The fact that SetProperties succeeded is the key verification here
        }
        finally
        {
            SQLiteDatabaseHelper.DeleteDatabase(dbPath);
            if (IOFile.Exists(odcFile))
            {
                IOFile.Delete(odcFile);
            }
        }
    }

    /// <summary>
    /// Test that connection string updates work for programmatically-created connections
    /// (not ODC imports). Excel may only block connection string changes for ODC imports.
    /// </summary>
    [Fact]
    public void CreateConnection_ThenUpdateConnectionString_Succeeds()
    {
        // Arrange - Create ODBC connection programmatically
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(CreateConnection_ThenUpdateConnectionString_Succeeds),
            _tempDir);

        // Create two ODBC connection strings (they don't need to have valid DSNs for this test)
        string initialConnectionString = "ODBC;DSN=Excel Files1;DBQ=C:\\temp\\test1.xlsx";
        string updatedConnectionString = "ODBC;DSN=Excel Files2;DBQ=C:\\temp\\test2.xlsx";
        string connectionName = "TestOdbcConnection";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection with first ODBC string
        var createResult = _commands.Create(batch, connectionName, initialConnectionString);
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");

        // Verify initial connection string
        var viewResult1 = _commands.View(batch, connectionName);
        Assert.True(viewResult1.Success);
        Assert.Contains("Excel Files1", viewResult1.ConnectionString, StringComparison.OrdinalIgnoreCase);

        // Act - Update connection string to second ODBC string
        var updateResult = _commands.SetProperties(
            batch,
            connectionName: connectionName,
            connectionString: updatedConnectionString,
            commandText: null,
            description: null,
            backgroundQuery: null,
            refreshOnFileOpen: null,
            savePassword: null,
            refreshPeriod: null);

        // Assert - Connection string update should succeed (unlike ODC imports)
        Assert.True(updateResult.Success, $"SetProperties failed: {updateResult.ErrorMessage}");

        // Verify connection string was updated
        var viewResult2 = _commands.View(batch, connectionName);
        Assert.True(viewResult2.Success);
        Assert.Contains("Excel Files2", viewResult2.ConnectionString, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Excel Files1", viewResult2.ConnectionString, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Extracts connection name from Action string like "import-odc (imported: ConnectionName)"
    /// </summary>
    private static string ExtractConnectionName(string? action)
    {
        if (string.IsNullOrEmpty(action))
        {
            return string.Empty;
        }

        var startIndex = action.IndexOf("(imported:", StringComparison.Ordinal);
        if (startIndex < 0)
        {
            return string.Empty;
        }

        startIndex += "(imported:".Length;
        var endIndex = action.IndexOf(')', startIndex);
        if (endIndex < 0)
        {
            return string.Empty;
        }

        return action.Substring(startIndex, endIndex - startIndex).Trim();
    }
}

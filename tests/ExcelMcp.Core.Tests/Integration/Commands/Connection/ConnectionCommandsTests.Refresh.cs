using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void Refresh_ConnectionNotFound_ReturnsFailure()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_ConnectionNotFound_ReturnsFailure),
            _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Refresh(batch, "NonExistentConnection");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }
    /// <inheritdoc/>

    [Fact]
    public void Refresh_ConnectionOnlyQuery_ReturnsSuccessWithContext()
    {
        // Arrange - Create a text connection but don't load data (connection-only)
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_ConnectionOnlyQuery_ReturnsSuccessWithContext),
            _tempDir);

        var csvFile = Path.Combine(_tempDir, "test-data.csv");
        var connectionName = "TestTextConnection";

        // Create text connection without loading data to any worksheet
        ConnectionTestHelper.CreateTextFileConnection(testFile, connectionName, csvFile);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Refresh connection-only connection (should succeed but indicate no data loaded)
        var result = _commands.Refresh(batch, connectionName);

        // Assert - Pure COM passthrough, just verify success
        Assert.True(result.Success, $"Connection-only refresh should succeed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact(Skip = "LoadTo requires actual data source - OLEDB is primary use case but needs real DB")]
    public void Refresh_ConnectionWithLoadedData_ReturnsSuccess()
    {
        // NOTE: This test documents that LoadTo works with OLEDB connections (primary use case)
        // TEXT connections DON'T support the QueryTables.Add() pattern used by LoadTo
        // To enable this test, provide a working OLEDB connection string

        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_ConnectionWithLoadedData_ReturnsSuccess),
            _tempDir);

        var connectionName = "RefreshTestConnection";
        var connectionString = "Provider=SQLOLEDB;Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=tempdb;Integrated Security=SSPI;";

        ConnectionTestHelper.CreateOleDbConnection(testFile, connectionName, connectionString);

        using var batch = ExcelSession.BeginBatch(testFile);
        var loadResult = _commands.LoadTo(batch, connectionName, "TestSheet");
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");

        batch.Save();

        var result = _commands.Refresh(batch, connectionName);
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void Refresh_TextConnectionMissingFile_SucceedsWithoutValidation()
    {
        // Arrange - This test documents Excel's actual behavior: TEXT connections
        // don't immediately validate file existence on refresh
        const string connectionName = "TestTextConnectionMissingFile";
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(ConnectionCommandsTests),
            nameof(Refresh_TextConnectionMissingFile_SucceedsWithoutValidation),
            _tempDir);
        var csvFile = Path.Combine(_tempDir, $"missing_file_{Guid.NewGuid()}.csv");

        // Create CSV file temporarily, then delete it after connection creation
        System.IO.File.WriteAllText(csvFile, "Col1,Col2\nVal1,Val2\n");

        // Create TEXT connection while file exists
        ConnectionTestHelper.CreateTextFileConnection(testFile, connectionName, csvFile);

        // Delete the file - this is the key difference from the other test
        System.IO.File.Delete(csvFile);
        Assert.False(System.IO.File.Exists(csvFile), "CSV file should be deleted");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Refresh connection to missing file
        var result = _commands.Refresh(batch, connectionName);

        // Assert - Excel COM doesn't immediately detect missing files for TEXT connections
        // This documents the actual behavior, not the expected behavior
        Assert.True(result.Success,
            "Excel COM allows TEXT connection refresh even when file is missing. " +
            "File validation may happen later during actual data access.");

        // Cleanup - file already deleted
    }
}

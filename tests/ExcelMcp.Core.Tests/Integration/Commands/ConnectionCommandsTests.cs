using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop.Session;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Comprehensive integration tests for ConnectionCommands
/// Tests all 11 connection operations with batch API pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class ConnectionCommandsTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFile;
    private readonly string _testCsvFile;
    private readonly ConnectionCommands _commands;
    private readonly FileCommands _fileCommands;

    public ConnectionCommandsTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_Conn_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _testFile = Path.Combine(_testDir, "test.xlsx");
        _testCsvFile = Path.Combine(_testDir, "data.csv");
        _commands = new ConnectionCommands();
        _fileCommands = new FileCommands();

        // Create test workbook
        var result = _fileCommands.CreateEmptyAsync(_testFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test file: {result.ErrorMessage}");
        }

        // Create test CSV file for text connections
        File.WriteAllText(_testCsvFile, "Name,Value\nTest1,100\nTest2,200");
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, recursive: true);
            }
        }
        catch { /* Cleanup failure is non-critical */ }
        GC.SuppressFinalize(this);
    }

    #region List Tests

    [Fact]
    public async Task List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Connections);
        Assert.Empty(result.Connections);
        Assert.Equal(_testFile, result.FilePath);
    }

    [Fact]
    public async Task List_WithOleDbConnection_ReturnsConnection()
    {
        // Arrange
        string connName = "TestOleDb";
        string connString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=TestDB";
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.Connections);
        var conn = Assert.Single(result.Connections);
        Assert.Equal(connName, conn.Name);
        Assert.Equal("OLEDB", conn.Type);
    }

    #endregion

    #region View Tests

    [Fact]
    public async Task View_ExistingConnection_ReturnsDetails()
    {
        // Arrange
        string connName = "TestConn";
        string connString = "Provider=SQLOLEDB;Data Source=localhost";
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ViewAsync(batch, connName);

        // Assert
        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(connName, result.ConnectionName);
        Assert.Equal("OLEDB", result.Type);
        Assert.NotNull(result.ConnectionString);
        Assert.Contains("SQLOLEDB", result.ConnectionString);
        Assert.NotNull(result.DefinitionJson);
    }

    [Fact]
    public async Task View_NonExistentConnection_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ViewAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task View_NullConnectionName_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ViewAsync(batch, null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Export Tests

    [Fact]
    public async Task Export_ExistingConnection_CreatesJsonFile()
    {
        // Arrange
        string connName = "ExportTest";
        string connString = "Provider=SQLOLEDB;Data Source=localhost";
        string jsonPath = Path.Combine(_testDir, "export.json");
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ExportAsync(batch, connName, jsonPath);

        // Assert
        Assert.True(result.Success, $"Export failed: {result.ErrorMessage}");
        Assert.True(File.Exists(jsonPath), "JSON file not created");

        // Verify JSON content
        string json = File.ReadAllText(jsonPath);
        Assert.NotEmpty(json);
        var jsonDoc = JsonDocument.Parse(json);
        Assert.Equal(connName, jsonDoc.RootElement.GetProperty("Name").GetString());
    }

    [Fact]
    public async Task Export_NonExistentConnection_ReturnsError()
    {
        // Arrange
        string jsonPath = Path.Combine(_testDir, "export.json");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ExportAsync(batch, "NonExistent", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.False(File.Exists(jsonPath), "JSON file should not be created on error");
    }

    #endregion

    #region Import Tests

    [Fact]
    public async Task Import_FromValidJson_CreatesConnection()
    {
        // Arrange
        string connName = "ImportedConn";
        string jsonPath = Path.Combine(_testDir, "import.json");

        // Create JSON definition
        var definition = new
        {
            Name = connName,
            Type = "OLEDB",
            Description = "Imported test connection",
            ConnectionString = "Provider=SQLOLEDB;Data Source=localhost",
            Properties = new { BackgroundQuery = true, RefreshOnFileOpen = false }
        };
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(definition, new JsonSerializerOptions { WriteIndented = true }));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ImportAsync(batch, connName, jsonPath);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Import failed: {result.ErrorMessage}");

        // Verify connection was created
        await using var verifyBatch = await ExcelSession.BeginBatchAsync(_testFile);
        var listResult = await _commands.ListAsync(verifyBatch);
        Assert.Contains(listResult.Connections, c => c.Name == connName);
    }

    [Fact]
    public async Task Import_InvalidJsonFile_ReturnsError()
    {
        // Arrange
        string jsonPath = Path.Combine(_testDir, "invalid.json");
        File.WriteAllText(jsonPath, "{ invalid json }");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ImportAsync(batch, "TestConn", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Update Tests

    [Fact]
    public async Task Update_ExistingConnection_ModifiesConnection()
    {
        // Arrange
        string connName = "UpdateTest";
        string connString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=DB1";
        string jsonPath = Path.Combine(_testDir, "update.json");

        // Create initial connection
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Create updated definition
        var updatedDefinition = new
        {
            Name = connName,
            Type = "OLEDB",
            Description = "Updated description",
            ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=DB2",
            Properties = new { BackgroundQuery = false }
        };
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(updatedDefinition, new JsonSerializerOptions { WriteIndented = true }));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.UpdateAsync(batch, connName, jsonPath);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Update failed: {result.ErrorMessage}");

        // Verify update
        await using var verifyBatch = await ExcelSession.BeginBatchAsync(_testFile);
        var viewResult = await _commands.ViewAsync(verifyBatch, connName);
        Assert.Contains("DB2", viewResult.ConnectionString);
    }

    [Fact]
    public async Task Update_NonExistentConnection_ReturnsError()
    {
        // Arrange
        string jsonPath = Path.Combine(_testDir, "update.json");
        var definition = new
        {
            Name = "NonExistent",
            Type = "OLEDB",
            ConnectionString = "Provider=SQLOLEDB;Data Source=localhost"
        };
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(definition));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.UpdateAsync(batch, "NonExistent", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Delete Tests

    [Fact]
    public async Task Delete_ExistingConnection_RemovesConnection()
    {
        // Arrange
        string connName = "DeleteTest";
        string connString = "Provider=SQLOLEDB;Data Source=localhost";
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.DeleteAsync(batch, connName);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify deletion
        await using var verifyBatch = await ExcelSession.BeginBatchAsync(_testFile);
        var listResult = await _commands.ListAsync(verifyBatch);
        Assert.DoesNotContain(listResult.Connections, c => c.Name == connName);
    }

    [Fact]
    public async Task Delete_NonExistentConnection_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.DeleteAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region GetProperties Tests

    [Fact]
    public async Task GetProperties_ExistingConnection_ReturnsProperties()
    {
        // Arrange
        string connName = "PropTest";
        string connString = "Provider=SQLOLEDB;Data Source=localhost";
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetPropertiesAsync(batch, connName);

        // Assert
        Assert.True(result.Success, $"GetProperties failed: {result.ErrorMessage}");
        Assert.Equal(connName, result.ConnectionName);
        // Properties are non-nullable bools, just verify result succeeded
    }

    [Fact]
    public async Task GetProperties_NonExistentConnection_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetPropertiesAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region SetProperties Tests

    [Fact]
    public async Task SetProperties_ExistingConnection_UpdatesProperties()
    {
        // Arrange
        string connName = "SetPropTest";
        string connString = "Provider=SQLOLEDB;Data Source=localhost";
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.SetPropertiesAsync(batch, connName,
            backgroundQuery: false,
            refreshOnFileOpen: true,
            savePassword: false,
            refreshPeriod: 60);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"SetProperties failed: {result.ErrorMessage}");

        // Verify properties were set
        await using var verifyBatch = await ExcelSession.BeginBatchAsync(_testFile);
        var propsResult = await _commands.GetPropertiesAsync(verifyBatch, connName);
        Assert.False(propsResult.BackgroundQuery);
        Assert.True(propsResult.RefreshOnFileOpen);
        Assert.False(propsResult.SavePassword);
        Assert.Equal(60, propsResult.RefreshPeriod);
    }

    [Fact]
    public async Task SetProperties_NonExistentConnection_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.SetPropertiesAsync(batch, "NonExistent", backgroundQuery: true);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Test Connection Tests

    [Fact]
    public async Task Test_ExistingConnection_ReturnsResult()
    {
        // Arrange
        string connName = "TestConnTest";
        string connString = "Provider=SQLOLEDB;Data Source=localhost";
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.TestAsync(batch, connName);

        // Assert
        Assert.NotNull(result);
        // Note: Test may fail or succeed depending on whether SQL Server is available
        // We just verify the method executes without throwing
    }

    [Fact]
    public async Task Test_NonExistentConnection_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.TestAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Refresh Tests

    [Fact]
    public async Task Refresh_NonExistentConnection_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.RefreshAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region LoadTo Tests

    [Fact]
    public async Task LoadTo_NonExistentConnection_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.LoadToAsync(batch, "NonExistent", "Sheet1");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Integration Workflow Tests

    [Fact]
    public async Task Workflow_ExportImportDelete_CompleteLifecycle()
    {
        // Step 1: Create connection
        string connName = "WorkflowTest";
        string connString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=TestDB";
        await ConnectionTestHelper.CreateOleDbConnectionAsync(_testFile, connName, connString);

        // Step 2: Export to JSON
        string jsonPath = Path.Combine(_testDir, "workflow.json");
        await using (var exportBatch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var exportResult = await _commands.ExportAsync(exportBatch, connName, jsonPath);
            Assert.True(exportResult.Success);
        }

        // Step 3: Delete original connection
        await using (var deleteBatch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var deleteResult = await _commands.DeleteAsync(deleteBatch, connName);
            await deleteBatch.SaveAsync();
            Assert.True(deleteResult.Success);
        }

        // Step 4: Verify deleted
        await using (var verifyBatch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var listResult = await _commands.ListAsync(verifyBatch);
            Assert.DoesNotContain(listResult.Connections, c => c.Name == connName);
        }

        // Step 5: Re-import from JSON
        await using (var importBatch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var importResult = await _commands.ImportAsync(importBatch, connName, jsonPath);
            await importBatch.SaveAsync();
            Assert.True(importResult.Success);
        }

        // Step 6: Verify re-imported
        await using (var finalBatch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var finalList = await _commands.ListAsync(finalBatch);
            Assert.Contains(finalList.Connections, c => c.Name == connName);
        }
    }

    #endregion
}

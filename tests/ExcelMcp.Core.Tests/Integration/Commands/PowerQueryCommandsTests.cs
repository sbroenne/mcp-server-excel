using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Power Query Core operations.
/// These tests require Excel installation and validate Core Power Query data operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
public class PowerQueryCommandsTests : IDisposable
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _testQueryFile;
    private readonly string _tempDir;
    private bool _disposed;

    public PowerQueryCommandsTests()
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_PQ_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");
        _testQueryFile = Path.Combine(_tempDir, "TestQuery.pq");

        // Create test Excel file and Power Query
        CreateTestExcelFile();
        CreateTestQueryFile();
    }

    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
    }

    private void CreateTestQueryFile()
    {
        // Create a simple Power Query M file that creates sample data
        // This avoids dependency on existing worksheets
        string mCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        File.WriteAllText(_testQueryFile, mCode);
    }

    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        Assert.Empty(result.Queries); // New file has no queries
    }

    [Fact]
    public async Task Import_WithValidMCode_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.ImportAsync(batch, "TestQuery", _testQueryFile);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    [Fact]
    public async Task List_AfterImport_ShowsNewQuery()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "TestQuery", _testQueryFile);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Queries);
        Assert.Single(result.Queries);
        Assert.Equal("TestQuery", result.Queries[0].Name);
        }
    }

    [Fact]
    public async Task View_WithExistingQuery_ReturnsMCode()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "TestQuery", _testQueryFile);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.ViewAsync(batch, "TestQuery");

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.MCode);
        Assert.Contains("Source", result.MCode);
        }
    }

    [Fact]
    public async Task Export_WithExistingQuery_CreatesFile()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "TestQuery", _testQueryFile);
            await batch.SaveAsync();
        }
        var exportPath = Path.Combine(_tempDir, "exported.pq");

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.ExportAsync(batch, "TestQuery", exportPath);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(exportPath));
        }
    }

    [Fact]
    public async Task Update_WithValidMCode_ReturnsSuccessResult()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "TestQuery", _testQueryFile);
            await batch.SaveAsync();
        }
        var updateFile = Path.Combine(_tempDir, "updated.pq");
        File.WriteAllText(updateFile, "let\n    UpdatedSource = 1\nin\n    UpdatedSource");

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.UpdateAsync(batch, "TestQuery", updateFile);
            await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
        }
    }

    [Fact]
    public async Task Delete_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "TestQuery", _testQueryFile);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.DeleteAsync(batch, "TestQuery");
            await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
        }
    }

    [Fact]
    public async Task View_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.ViewAsync(batch, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task Import_ThenDelete_ThenList_ShowsEmpty()
    {
        // Arrange & Act - Import
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "TestQuery", _testQueryFile);
            await batch.SaveAsync();
        }

        // Act - Delete
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.DeleteAsync(batch, "TestQuery");
            await batch.SaveAsync();
        }

        // Act - List
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success);
        Assert.Empty(result.Queries);
        }
    }

    [Fact]
    public async Task SetConnectionOnly_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestConnectionOnly", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "TestConnectionOnly");
            await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"SetConnectionOnly failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-connection-only", result.Action);
        }
    }

    [Fact]
    public async Task SetLoadToTable_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToTable", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.SetLoadToTableAsync(batch, "TestLoadToTable", "TestSheet");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success, $"SetLoadToTable failed: {result.ErrorMessage}");
            Assert.Equal("pq-set-load-to-table", result.Action);
        }
    }

    [Fact]
    public async Task SetLoadToDataModel_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToDataModel", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "TestLoadToDataModel");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success, $"SetLoadToDataModel failed: {result.ErrorMessage}");
            Assert.Equal("pq-set-load-to-data-model", result.Action);
        }
    }

    [Fact]
    public async Task SetLoadToBoth_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToBoth", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.SetLoadToBothAsync(batch, "TestLoadToBoth", "TestSheet");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success, $"SetLoadToBoth failed: {result.ErrorMessage}");
            Assert.Equal("pq-set-load-to-both", result.Action);
        }
    }

    [Fact]
    public async Task GetLoadConfig_WithConnectionOnlyQuery_ReturnsConnectionOnlyMode()
    {
        // Arrange - Import and set to connection only
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestConnectionOnlyConfig", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setResult = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "TestConnectionOnlyConfig");
            await batch.SaveAsync();
            Assert.True(setResult.Success, $"Failed to set connection only: {setResult.ErrorMessage}");
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestConnectionOnlyConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestConnectionOnlyConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, result.LoadMode);
        Assert.Null(result.TargetSheet);
        Assert.False(result.IsLoadedToDataModel);
        }
    }

    [Fact]
    public async Task GetLoadConfig_WithLoadToTableQuery_ReturnsLoadToTableMode()
    {
        // Arrange - Import and set to load to table
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToTableConfig", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setResult = await _powerQueryCommands.SetLoadToTableAsync(batch, "TestLoadToTableConfig", "ConfigTestSheet");
            await batch.SaveAsync();
            Assert.True(setResult.Success, $"Failed to set load to table: {setResult.ErrorMessage}");
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestLoadToTableConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestLoadToTableConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, result.LoadMode);
        Assert.Equal("ConfigTestSheet", result.TargetSheet);
        Assert.False(result.IsLoadedToDataModel);
        }
    }

    [Fact]
    public async Task GetLoadConfig_WithLoadToDataModelQuery_ReturnsLoadToDataModelMode()
    {
        // Arrange - Import and set to load to data model
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToDataModelConfig", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setResult = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "TestLoadToDataModelConfig");
            await batch.SaveAsync();
            Assert.True(setResult.Success, $"Failed to set load to data model: {setResult.ErrorMessage}");

            // Debug output
            if (!string.IsNullOrEmpty(setResult.ErrorMessage))
            {
                System.Console.WriteLine($"SetLoadToDataModel message: {setResult.ErrorMessage}");
            }
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestLoadToDataModelConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestLoadToDataModelConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, result.LoadMode);
        Assert.Null(result.TargetSheet);
        Assert.True(result.IsLoadedToDataModel);
        }
    }

    [Fact]
    public async Task GetLoadConfig_WithLoadToBothQuery_ReturnsLoadToBothMode()
    {
        // Arrange - Import and set to load to both
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToBothConfig", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setResult = await _powerQueryCommands.SetLoadToBothAsync(batch, "TestLoadToBothConfig", "BothTestSheet");
            await batch.SaveAsync();
            Assert.True(setResult.Success, $"Failed to set load to both: {setResult.ErrorMessage}");
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestLoadToBothConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestLoadToBothConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, result.LoadMode);
        Assert.Equal("BothTestSheet", result.TargetSheet);
        Assert.True(result.IsLoadedToDataModel);
        }
    }

    [Fact]
    public async Task LoadConfigurationWorkflow_SwitchingModes_UpdatesCorrectly()
    {
        // Arrange - Import a query
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var importResult = await _powerQueryCommands.ImportAsync(batch, "TestWorkflowQuery", _testQueryFile);
            await batch.SaveAsync();
            Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");
        }

        // Act & Assert - Test switching between different load modes

        // 1. Set to Connection Only
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setConnectionOnlyResult = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "TestWorkflowQuery");
            await batch.SaveAsync();
            Assert.True(setConnectionOnlyResult.Success, $"SetConnectionOnly failed: {setConnectionOnlyResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var getConnectionOnlyResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestWorkflowQuery");
            Assert.True(getConnectionOnlyResult.Success, $"GetLoadConfig after SetConnectionOnly failed: {getConnectionOnlyResult.ErrorMessage}");
            Assert.Equal(PowerQueryLoadMode.ConnectionOnly, getConnectionOnlyResult.LoadMode);
        }

        // 2. Switch to Load to Table
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setLoadToTableResult = await _powerQueryCommands.SetLoadToTableAsync(batch, "TestWorkflowQuery", "WorkflowSheet");
            await batch.SaveAsync();
            Assert.True(setLoadToTableResult.Success, $"SetLoadToTable failed: {setLoadToTableResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var getLoadToTableResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestWorkflowQuery");
            Assert.True(getLoadToTableResult.Success, $"GetLoadConfig after SetLoadToTable failed: {getLoadToTableResult.ErrorMessage}");
            Assert.Equal(PowerQueryLoadMode.LoadToTable, getLoadToTableResult.LoadMode);
            Assert.Equal("WorkflowSheet", getLoadToTableResult.TargetSheet);
        }

        // 3. Switch to Load to Data Model
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setLoadToDataModelResult = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "TestWorkflowQuery");
            await batch.SaveAsync();
            Assert.True(setLoadToDataModelResult.Success, $"SetLoadToDataModel failed: {setLoadToDataModelResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var getLoadToDataModelResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestWorkflowQuery");
            Assert.True(getLoadToDataModelResult.Success, $"GetLoadConfig after SetLoadToDataModel failed: {getLoadToDataModelResult.ErrorMessage}");
            Assert.Equal(PowerQueryLoadMode.LoadToDataModel, getLoadToDataModelResult.LoadMode);
            Assert.True(getLoadToDataModelResult.IsLoadedToDataModel);
        }

        // 4. Switch to Load to Both
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setLoadToBothResult = await _powerQueryCommands.SetLoadToBothAsync(batch, "TestWorkflowQuery", "BothWorkflowSheet");
            await batch.SaveAsync();
            Assert.True(setLoadToBothResult.Success, $"SetLoadToBoth failed: {setLoadToBothResult.ErrorMessage}");
        }

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var getLoadToBothResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestWorkflowQuery");
            Assert.True(getLoadToBothResult.Success, $"GetLoadConfig after SetLoadToBoth failed: {getLoadToBothResult.ErrorMessage}");
            Assert.Equal(PowerQueryLoadMode.LoadToBoth, getLoadToBothResult.LoadMode);
            Assert.Equal("BothWorkflowSheet", getLoadToBothResult.TargetSheet);
            Assert.True(getLoadToBothResult.IsLoadedToDataModel);
        }
    }

    [Fact]
    public async Task GetLoadConfig_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.GetLoadConfigAsync(batch, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
        Assert.Equal("NonExistentQuery", result.QueryName);
    }

    [Fact]
    public async Task SetLoadToTable_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.SetLoadToTableAsync(batch, "NonExistentQuery", "TestSheet");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public async Task SetLoadToDataModel_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public async Task SetLoadToBoth_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.SetLoadToBothAsync(batch, "NonExistentQuery", "TestSheet");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public async Task SetConnectionOnly_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, true);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        _disposed = true;
    }
}

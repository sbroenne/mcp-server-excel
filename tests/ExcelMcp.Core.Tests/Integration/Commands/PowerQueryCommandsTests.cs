using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
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
public class CorePowerQueryCommandsTests : IDisposable
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _testQueryFile;
    private readonly string _tempDir;
    private bool _disposed;

    public CorePowerQueryCommandsTests()
    {
        _powerQueryCommands = new PowerQueryCommands();
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
        var result = _fileCommands.CreateEmpty(_testExcelFile, overwriteIfExists: false);
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
    public void List_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _powerQueryCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        Assert.Empty(result.Queries); // New file has no queries
    }

    [Fact]
    public async Task Import_WithValidMCode_ReturnsSuccessResult()
    {
        // Act
        var result = await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    [Fact]
    public async Task List_AfterImport_ShowsNewQuery()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Act
        var result = _powerQueryCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Queries);
        Assert.Single(result.Queries);
        Assert.Equal("TestQuery", result.Queries[0].Name);
    }

    [Fact]
    public async Task View_WithExistingQuery_ReturnsMCode()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Act
        var result = _powerQueryCommands.View(_testExcelFile, "TestQuery");

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.MCode);
        Assert.Contains("Source", result.MCode);
    }

    [Fact]
    public async Task Export_WithExistingQuery_CreatesFile()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);
        var exportPath = Path.Combine(_tempDir, "exported.pq");

        // Act
        var result = await _powerQueryCommands.Export(_testExcelFile, "TestQuery", exportPath);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public async Task Update_WithValidMCode_ReturnsSuccessResult()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);
        var updateFile = Path.Combine(_tempDir, "updated.pq");
        File.WriteAllText(updateFile, "let\n    UpdatedSource = 1\nin\n    UpdatedSource");

        // Act
        var result = await _powerQueryCommands.Update(_testExcelFile, "TestQuery", updateFile);

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task Delete_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);

        // Act
        var result = _powerQueryCommands.Delete(_testExcelFile, "TestQuery");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void View_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        var result = _powerQueryCommands.View(_testExcelFile, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task Import_ThenDelete_ThenList_ShowsEmpty()
    {
        // Arrange
        await _powerQueryCommands.Import(_testExcelFile, "TestQuery", _testQueryFile);
        _powerQueryCommands.Delete(_testExcelFile, "TestQuery");

        // Act
        var result = _powerQueryCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.Empty(result.Queries);
    }

    [Fact]
    public async Task SetConnectionOnly_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestConnectionOnly", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act
        var result = _powerQueryCommands.SetConnectionOnly(_testExcelFile, "TestConnectionOnly");

        // Assert
        Assert.True(result.Success, $"SetConnectionOnly failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-connection-only", result.Action);
    }

    [Fact]
    public async Task SetLoadToTable_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestLoadToTable", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act
        var result = _powerQueryCommands.SetLoadToTable(_testExcelFile, "TestLoadToTable", "TestSheet");

        // Assert
        Assert.True(result.Success, $"SetLoadToTable failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-table", result.Action);
    }

    [Fact]
    public async Task SetLoadToDataModel_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestLoadToDataModel", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act
        var result = _powerQueryCommands.SetLoadToDataModel(_testExcelFile, "TestLoadToDataModel");

        // Assert
        Assert.True(result.Success, $"SetLoadToDataModel failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-data-model", result.Action);
    }

    [Fact]
    public async Task SetLoadToBoth_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Import a query first
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestLoadToBoth", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act
        var result = _powerQueryCommands.SetLoadToBoth(_testExcelFile, "TestLoadToBoth", "TestSheet");

        // Assert
        Assert.True(result.Success, $"SetLoadToBoth failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-both", result.Action);
    }

    [Fact]
    public async Task GetLoadConfig_WithConnectionOnlyQuery_ReturnsConnectionOnlyMode()
    {
        // Arrange - Import and set to connection only
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestConnectionOnlyConfig", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var setResult = _powerQueryCommands.SetConnectionOnly(_testExcelFile, "TestConnectionOnlyConfig");
        Assert.True(setResult.Success, $"Failed to set connection only: {setResult.ErrorMessage}");

        // Act
        var result = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestConnectionOnlyConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestConnectionOnlyConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, result.LoadMode);
        Assert.Null(result.TargetSheet);
        Assert.False(result.IsLoadedToDataModel);
    }

    [Fact]
    public async Task GetLoadConfig_WithLoadToTableQuery_ReturnsLoadToTableMode()
    {
        // Arrange - Import and set to load to table
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestLoadToTableConfig", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var setResult = _powerQueryCommands.SetLoadToTable(_testExcelFile, "TestLoadToTableConfig", "ConfigTestSheet");
        Assert.True(setResult.Success, $"Failed to set load to table: {setResult.ErrorMessage}");

        // Act
        var result = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestLoadToTableConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestLoadToTableConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, result.LoadMode);
        Assert.Equal("ConfigTestSheet", result.TargetSheet);
        Assert.False(result.IsLoadedToDataModel);
    }

    [Fact]
    public async Task GetLoadConfig_WithLoadToDataModelQuery_ReturnsLoadToDataModelMode()
    {
        // Arrange - Import and set to load to data model
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestLoadToDataModelConfig", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var setResult = _powerQueryCommands.SetLoadToDataModel(_testExcelFile, "TestLoadToDataModelConfig");
        Assert.True(setResult.Success, $"Failed to set load to data model: {setResult.ErrorMessage}");

        // Debug output
        if (!string.IsNullOrEmpty(setResult.ErrorMessage))
        {
            System.Console.WriteLine($"SetLoadToDataModel message: {setResult.ErrorMessage}");
        }

        // Act
        var result = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestLoadToDataModelConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestLoadToDataModelConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, result.LoadMode);
        Assert.Null(result.TargetSheet);
        Assert.True(result.IsLoadedToDataModel);
    }

    [Fact]
    public async Task GetLoadConfig_WithLoadToBothQuery_ReturnsLoadToBothMode()
    {
        // Arrange - Import and set to load to both
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestLoadToBothConfig", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var setResult = _powerQueryCommands.SetLoadToBoth(_testExcelFile, "TestLoadToBothConfig", "BothTestSheet");
        Assert.True(setResult.Success, $"Failed to set load to both: {setResult.ErrorMessage}");

        // Act
        var result = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestLoadToBothConfig");

        // Assert
        Assert.True(result.Success, $"GetLoadConfig failed: {result.ErrorMessage}");
        Assert.Equal("TestLoadToBothConfig", result.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, result.LoadMode);
        Assert.Equal("BothTestSheet", result.TargetSheet);
        Assert.True(result.IsLoadedToDataModel);
    }

    [Fact]
    public async Task LoadConfigurationWorkflow_SwitchingModes_UpdatesCorrectly()
    {
        // Arrange - Import a query
        var importResult = await _powerQueryCommands.Import(_testExcelFile, "TestWorkflowQuery", _testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act & Assert - Test switching between different load modes

        // 1. Set to Connection Only
        var setConnectionOnlyResult = _powerQueryCommands.SetConnectionOnly(_testExcelFile, "TestWorkflowQuery");
        Assert.True(setConnectionOnlyResult.Success, $"SetConnectionOnly failed: {setConnectionOnlyResult.ErrorMessage}");

        var getConnectionOnlyResult = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestWorkflowQuery");
        Assert.True(getConnectionOnlyResult.Success, $"GetLoadConfig after SetConnectionOnly failed: {getConnectionOnlyResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, getConnectionOnlyResult.LoadMode);

        // 2. Switch to Load to Table
        var setLoadToTableResult = _powerQueryCommands.SetLoadToTable(_testExcelFile, "TestWorkflowQuery", "WorkflowSheet");
        Assert.True(setLoadToTableResult.Success, $"SetLoadToTable failed: {setLoadToTableResult.ErrorMessage}");

        var getLoadToTableResult = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestWorkflowQuery");
        Assert.True(getLoadToTableResult.Success, $"GetLoadConfig after SetLoadToTable failed: {getLoadToTableResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, getLoadToTableResult.LoadMode);
        Assert.Equal("WorkflowSheet", getLoadToTableResult.TargetSheet);

        // 3. Switch to Load to Data Model
        var setLoadToDataModelResult = _powerQueryCommands.SetLoadToDataModel(_testExcelFile, "TestWorkflowQuery");
        Assert.True(setLoadToDataModelResult.Success, $"SetLoadToDataModel failed: {setLoadToDataModelResult.ErrorMessage}");

        var getLoadToDataModelResult = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestWorkflowQuery");
        Assert.True(getLoadToDataModelResult.Success, $"GetLoadConfig after SetLoadToDataModel failed: {getLoadToDataModelResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, getLoadToDataModelResult.LoadMode);
        Assert.True(getLoadToDataModelResult.IsLoadedToDataModel);

        // 4. Switch to Load to Both
        var setLoadToBothResult = _powerQueryCommands.SetLoadToBoth(_testExcelFile, "TestWorkflowQuery", "BothWorkflowSheet");
        Assert.True(setLoadToBothResult.Success, $"SetLoadToBoth failed: {setLoadToBothResult.ErrorMessage}");

        var getLoadToBothResult = _powerQueryCommands.GetLoadConfig(_testExcelFile, "TestWorkflowQuery");
        Assert.True(getLoadToBothResult.Success, $"GetLoadConfig after SetLoadToBoth failed: {getLoadToBothResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, getLoadToBothResult.LoadMode);
        Assert.Equal("BothWorkflowSheet", getLoadToBothResult.TargetSheet);
        Assert.True(getLoadToBothResult.IsLoadedToDataModel);
    }

    [Fact]
    public void GetLoadConfig_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        var result = _powerQueryCommands.GetLoadConfig(_testExcelFile, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
        Assert.Equal("NonExistentQuery", result.QueryName);
    }

    [Fact]
    public void SetLoadToTable_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        var result = _powerQueryCommands.SetLoadToTable(_testExcelFile, "NonExistentQuery", "TestSheet");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public void SetLoadToDataModel_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        var result = _powerQueryCommands.SetLoadToDataModel(_testExcelFile, "NonExistentQuery");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public void SetLoadToBoth_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        var result = _powerQueryCommands.SetLoadToBoth(_testExcelFile, "NonExistentQuery", "TestSheet");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public void SetConnectionOnly_WithNonExistentQuery_ReturnsErrorResult()
    {
        // Act
        var result = _powerQueryCommands.SetConnectionOnly(_testExcelFile, "NonExistentQuery");

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

using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Power Query load configuration workflow.
/// Tests the complete workflow of switching between different load modes.
/// This comprehensive test validates mode switching and verification.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
public class PowerQueryLoadConfigWorkflowTests : IDisposable
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _testQueryFile;
    private readonly string _tempDir;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the test class.
    /// Creates a temporary directory and test Excel file with a sample Power Query.
    /// </summary>
    public PowerQueryLoadConfigWorkflowTests()
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_PQ_Workflow_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");
        _testQueryFile = Path.Combine(_tempDir, "TestQuery.pq");

        // Create test Excel file and Power Query
        CreateTestExcelFile();
        CreateTestQueryFile();
    }

    /// <summary>
    /// Creates an empty Excel workbook file for testing.
    /// Throws InvalidOperationException if Excel is not installed.
    /// </summary>
    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
    }

    /// <summary>
    /// Creates a test Power Query M code file with sample table data.
    /// The query creates a simple 3x3 table without external dependencies.
    /// </summary>
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

    /// <summary>
    /// Verifies the complete load configuration workflow by switching between all modes.
    /// Tests transitions: ConnectionOnly → LoadToTable → LoadToDataModel → LoadToBoth.
    /// Validates that each mode change is correctly applied and persisted.
    /// This comprehensive test validates mode switching and verification, eliminating the need
    /// for separate tests for each individual Get/Set operation pair.
    /// </summary>
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

    /// <summary>
    /// Disposes test resources and cleans up temporary directory.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Protected implementation of Dispose pattern.
    /// Cleans up temporary test directory if it exists.
    /// </summary>
    /// <param name="disposing">True if called from Dispose(), false if called from finalizer</param>
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

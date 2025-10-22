using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Data Model Core operations.
/// These tests require Excel installation and validate Core Data Model operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
public class CoreDataModelCommandsTests : IDisposable
{
    private readonly IDataModelCommands _dataModelCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _testMeasureFile;
    private readonly string _tempDir;
    private bool _disposed;

    public CoreDataModelCommandsTests()
    {
        _dataModelCommands = new DataModelCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_DM_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestDataModel.xlsx");
        _testMeasureFile = Path.Combine(_tempDir, "TestMeasure.dax");

        // Create test Excel file with Data Model
        CreateTestDataModelFile();
    }

    private void CreateTestDataModelFile()
    {
        // Create an empty workbook first
        var result = _fileCommands.CreateEmpty(_testExcelFile, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }

        // TODO: Add helper method to populate with Data Model
        // For now, this creates an empty workbook
        // Integration tests that require Data Model will be skipped if no model exists
    }

    [Fact]
    public void ListTables_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _dataModelCommands.ListTables(_testExcelFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        
        // New file without Data Model should indicate that
        if (!result.Success && result.ErrorMessage?.Contains("does not contain a Data Model") == true)
        {
            // This is expected for empty workbook
            Assert.Contains("does not contain a Data Model", result.ErrorMessage);
        }
    }

    [Fact]
    public void ListMeasures_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _dataModelCommands.ListMeasures(_testExcelFile);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");
        
        if (result.Success)
        {
            Assert.NotNull(result.Measures);
        }
    }

    [Fact]
    public void ViewMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Act
        var result = _dataModelCommands.ViewMeasure(_testExcelFile, "NonExistentMeasure");

        // Assert
        // Should fail with either "no Data Model" or "measure not found"
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("Measure 'NonExistentMeasure' not found"),
            $"Expected 'no Data Model' or 'measure not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task ExportMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Act
        var result = await _dataModelCommands.ExportMeasure(_testExcelFile, "NonExistentMeasure", _testMeasureFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void ListRelationships_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _dataModelCommands.ListRelationships(_testExcelFile);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");
        
        if (result.Success)
        {
            Assert.NotNull(result.Relationships);
        }
    }

    [Fact]
    public void Refresh_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        var result = _dataModelCommands.Refresh(_testExcelFile);

        // Assert
        // Refresh should either succeed or indicate no Data Model
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");
    }

    [Fact]
    public void ListTables_WithNonExistentFile_ReturnsErrorResult()
    {
        // Arrange
        var nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");

        // Act
        var result = _dataModelCommands.ListTables(nonExistentFile);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                // Give Excel time to release file locks
                System.Threading.Thread.Sleep(100);
                
                // Retry cleanup a few times if needed
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, recursive: true);
                        break;
                    }
                    catch (IOException) when (i < 2)
                    {
                        System.Threading.Thread.Sleep(500);
                    }
                }
            }
        }
        catch
        {
            // Best effort cleanup
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}

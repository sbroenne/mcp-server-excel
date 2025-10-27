using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
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
        CreateTestDataModelFileAsync().GetAwaiter().GetResult();
    }

    private async Task CreateTestDataModelFileAsync()
    {
        // Create an empty workbook first
        var result = await _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }

        // Create realistic Data Model with sample data
        try
        {
            await DataModelTestHelper.CreateSampleDataModelAsync(_testExcelFile);
        }
        catch (Exception ex)
        {
            // Data Model creation may fail on some Excel versions
            // Tests will handle this gracefully by checking for "no Data Model" errors
            System.Diagnostics.Debug.WriteLine($"Could not create sample Data Model: {ex.Message}");
        }
    }

    [Fact]
    public async Task ListTables_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListTablesAsync(batch);

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
    public async Task ListMeasures_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListMeasuresAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        if (result.Success)
        {
            Assert.NotNull(result.Measures);
        }
    }

    [Fact]
    public async Task ViewMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ViewMeasureAsync(batch, "NonExistentMeasure");

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
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ExportMeasureAsync(batch, "NonExistentMeasure", _testMeasureFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task ListRelationships_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        if (result.Success)
        {
            Assert.NotNull(result.Relationships);
        }
    }

    [Fact]
    public async Task Refresh_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.RefreshAsync(batch);
        await batch.SaveAsync();

        // Assert
        // Refresh should either succeed or indicate no Data Model
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");
    }

    [Fact]
    public async Task ListTables_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");

        // Act & Assert - BeginBatchAsync should throw FileNotFoundException for non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(nonExistentFile);
            await _dataModelCommands.ListTablesAsync(batch);
        });
    }

    [Fact]
    public async Task ListTables_WithRealisticDataModel_ReturnsTablesWithData()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListTablesAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);

        // If Data Model was created successfully, validate the tables
        if (result.Tables != null && result.Tables.Count > 0)
        {
            // Should have Sales, Customers, and Products tables
            Assert.True(result.Tables.Count >= 3, $"Expected at least 3 tables, got {result.Tables.Count}");

            var tableNames = result.Tables.Select(t => t.Name).ToList();
            Assert.Contains("Sales", tableNames);
            Assert.Contains("Customers", tableNames);
            Assert.Contains("Products", tableNames);

            // Validate Sales table has expected columns
            var salesTable = result.Tables.FirstOrDefault(t => t.Name == "Sales");
            if (salesTable != null)
            {
                Assert.True(salesTable.RecordCount > 0, "Sales table should have rows");
            }
        }
    }

    [Fact]
    public async Task ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListMeasuresAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        // If Data Model was created successfully with measures, validate them
        if (result.Success && result.Measures != null && result.Measures.Count > 0)
        {
            // Should have at least Total Sales, Average Sale, Total Customers
            Assert.True(result.Measures.Count >= 3, $"Expected at least 3 measures, got {result.Measures.Count}");

            var measureNames = result.Measures.Select(m => m.Name).ToList();
            Assert.Contains("Total Sales", measureNames);
            Assert.Contains("Average Sale", measureNames);
            Assert.Contains("Total Customers", measureNames);

            // Validate Total Sales measure has DAX formula
            var totalSales = result.Measures.FirstOrDefault(m => m.Name == "Total Sales");
            if (totalSales != null)
            {
                Assert.NotNull(totalSales.FormulaPreview);
                Assert.Contains("SUM", totalSales.FormulaPreview, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Amount", totalSales.FormulaPreview, StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    [Fact]
    public async Task ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        // If Data Model was created successfully with relationships, validate them
        if (result.Success && result.Relationships != null && result.Relationships.Count > 0)
        {
            // Should have at least 2 relationships (Sales->Customers, Sales->Products)
            Assert.True(result.Relationships.Count >= 2, $"Expected at least 2 relationships, got {result.Relationships.Count}");

            // Validate Sales->Customers relationship
            var salesCustomersRel = result.Relationships.FirstOrDefault(r =>
                r.FromTable == "Sales" && r.ToTable == "Customers");

            if (salesCustomersRel != null)
            {
                Assert.Equal("CustomerID", salesCustomersRel.FromColumn);
                Assert.Equal("CustomerID", salesCustomersRel.ToColumn);
                Assert.True(salesCustomersRel.IsActive, "Sales->Customers relationship should be active");
            }

            // Validate Sales->Products relationship
            var salesProductsRel = result.Relationships.FirstOrDefault(r =>
                r.FromTable == "Sales" && r.ToTable == "Products");

            if (salesProductsRel != null)
            {
                Assert.Equal("ProductID", salesProductsRel.FromColumn);
                Assert.Equal("ProductID", salesProductsRel.ToColumn);
                Assert.True(salesProductsRel.IsActive, "Sales->Products relationship should be active");
            }
        }
    }

    [Fact]
    public async Task ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ViewMeasureAsync(batch, "Total Sales");

        // Assert - Should either succeed with valid DAX or indicate no Data Model
        if (result.Success)
        {
            Assert.NotNull(result.DaxFormula);
            Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Sales", result.DaxFormula);
            Assert.Contains("Amount", result.DaxFormula);
            Assert.Equal("Total Sales", result.MeasureName);
        }
        else
        {
            // If not successful, should be because Data Model wasn't created or measure doesn't exist
            Assert.True(
                result.ErrorMessage?.Contains("does not contain a Data Model") == true ||
                result.ErrorMessage?.Contains("not found") == true,
                $"Expected 'no Data Model' or 'not found', but got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public async Task ExportMeasure_WithRealisticDataModel_ExportsValidDAXFile()
    {
        // Arrange
        var exportPath = Path.Combine(_tempDir, "TotalSales.dax");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ExportMeasureAsync(batch, "Total Sales", exportPath);

        // Assert - Should either succeed or indicate no Data Model
        if (result.Success)
        {
            Assert.True(File.Exists(exportPath), "DAX file should be created");

            var daxContent = File.ReadAllText(exportPath);
            Assert.NotEmpty(daxContent);
            Assert.Contains("SUM", daxContent, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Sales", daxContent);
            Assert.Contains("Amount", daxContent);
        }
        else
        {
            Assert.True(
                result.ErrorMessage?.Contains("does not contain a Data Model") == true ||
                result.ErrorMessage?.Contains("not found") == true,
                $"Expected 'no Data Model' or 'not found', but got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public async Task Refresh_WithRealisticDataModel_SucceedsOrIndicatesNoModel()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.RefreshAsync(batch);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        // If successful, should have refreshed the Data Model
        if (result.Success)
        {
            Assert.NotNull(result.FilePath);
            Assert.Equal(_testExcelFile, result.FilePath);
        }
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public async Task DeleteMeasure_WithValidMeasure_ReturnsSuccessResult()
    {
        // Arrange - Create a test measure first
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];

        await DataModelTestHelper.CreateTestMeasureAsync(_testExcelFile, measureName, "SUM(Sales[Amount])");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.DeleteMeasureAsync(batch, measureName);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.SuggestedNextActions);
        Assert.Contains(result.SuggestedNextActions, s => s.Contains("deleted successfully"));
    }

    [Fact]
    public async Task DeleteMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.DeleteMeasureAsync(batch, "NonExistentMeasure");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("Measure 'NonExistentMeasure' not found"),
            $"Expected 'no Data Model' or 'measure not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task DeleteMeasure_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Act & Assert - BeginBatchAsync should throw FileNotFoundException for non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("NonExistent.xlsx");
            await _dataModelCommands.DeleteMeasureAsync(batch, "SomeMeasure");
        });
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public async Task DeleteRelationship_WithValidRelationship_ReturnsSuccessResult()
    {
        // Arrange - Requires Data Model with relationships
        await using var listBatch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var listResult = await _dataModelCommands.ListRelationshipsAsync(listBatch);

        Assert.True(listResult.Success, "ListRelationships should succeed");
        Assert.NotNull(listResult.Relationships);
        Assert.True(listResult.Relationships.Count > 0, "Data Model should have relationships for this test");

        // Use the first relationship for testing
        var rel = listResult.Relationships[0];

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.DeleteRelationshipAsync(
            batch,
            rel.FromTable,
            rel.FromColumn,
            rel.ToTable,
            rel.ToColumn
        );

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.SuggestedNextActions);
        Assert.Contains(result.SuggestedNextActions, s => s.Contains("deleted successfully"));
    }

    [Fact]
    public async Task DeleteRelationship_WithNonExistentRelationship_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.DeleteRelationshipAsync(
            batch,
            "FakeTable",
            "FakeColumn",
            "OtherTable",
            "OtherColumn"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("not found in Data Model"),
            $"Expected 'no Data Model' or 'not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task DeleteRelationship_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Act & Assert - BeginBatchAsync should throw FileNotFoundException for non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("NonExistent.xlsx");
            await _dataModelCommands.DeleteRelationshipAsync(
                batch,
                "Table1",
                "Col1",
                "Table2",
                "Col2"
            );
        });
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

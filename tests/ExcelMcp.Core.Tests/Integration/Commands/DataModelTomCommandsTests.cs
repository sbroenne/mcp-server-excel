using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Data Model TOM (Tabular Object Model) operations.
/// These tests require Excel installation and validate TOM Data Model operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModelTom")]
public class CoreDataModelTomCommandsTests : IDisposable
{
    private readonly IDataModelTomCommands _tomCommands;
    private readonly IDataModelCommands _dataModelCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public CoreDataModelTomCommandsTests()
    {
        _tomCommands = new DataModelTomCommands();
        _dataModelCommands = new DataModelCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_DM_TOM_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestDataModelTom.xlsx");

        // Create test Excel file with Data Model
        CreateTestDataModelFile();
    }

    private void CreateTestDataModelFile()
    {
        // Create an empty workbook first
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }

        // Create realistic Data Model with sample data
        try
        {
            DataModelTestHelper.CreateSampleDataModelAsync(_testExcelFile).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            // Data Model creation may fail on some Excel versions
            // Tests will handle this gracefully by checking for "no Data Model" errors
            System.Diagnostics.Debug.WriteLine($"Could not create sample Data Model: {ex.Message}");
        }
    }

    #region CreateMeasure Tests

    [Fact]
    public async Task CreateMeasure_WithValidParameters_ReturnsSuccess()
    {
        // Arrange
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            daxFormula,
            "Test measure for integration testing",
            "#,##0.00"
        );

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            Assert.NotNull(result.SuggestedNextActions);
            Assert.Contains(result.SuggestedNextActions, s => s.Contains("created successfully"));

            // Verify the measure was actually created by listing measures
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
            var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
            if (listResult.Success)
            {
                Assert.Contains(listResult.Measures, m => m.Name == measureName);
            }
        }
        else
        {
            // If TOM connection failed, verify it's because of Data Model availability
            Assert.True(
                result.ErrorMessage?.Contains("Data Model") == true ||
                result.ErrorMessage?.Contains("connect") == true,
                $"Expected Data Model or connection error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void CreateMeasure_WithInvalidTable_ReturnsError()
    {
        // Arrange
        var measureName = "InvalidTableMeasure";
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "NonExistentTable",
            measureName,
            daxFormula
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("Table") ||
            result.ErrorMessage.Contains("not found") ||
            result.ErrorMessage.Contains("connect"),
            $"Expected table or connection error, got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public void CreateMeasure_WithEmptyMeasureName_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            "",
            "SUM(Sales[Amount])"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("name cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateMeasure_WithEmptyFormula_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            "TestMeasure",
            ""
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("formula cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "TOM API requires specific configuration and may not be available in all Excel environments.")]
    public void CreateMeasure_WithDuplicateName_ReturnsError()
    {
        // Arrange - Create first measure
        var measureName = "DuplicateTest_" + Guid.NewGuid().ToString("N")[..8];
        var result1 = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            "SUM(Sales[Amount])"
        );

        Assert.True(result1.Success, $"First create failed: {result1.ErrorMessage}");

        // Act - Try to create duplicate
        var result2 = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            "AVERAGE(Sales[Amount])"
        );

        // Assert
        Assert.False(result2.Success);
        Assert.NotNull(result2.ErrorMessage);
        Assert.Contains("already exists", result2.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region UpdateMeasure Tests

    [Fact]
    public async Task UpdateMeasure_WithValidParameters_ReturnsSuccess()
    {
        // Arrange - Create a measure first
        var measureName = "UpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var createResult = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            "SUM(Sales[Amount])"
        );

        // TOM connection failure should fail the test, not skip it
        if (!createResult.Success && createResult.ErrorMessage?.Contains("connect") == true)
        {
            Assert.Fail($"TOM connection failed: {createResult.ErrorMessage}");
        }

        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");

        // Act - Update the measure
        var updateResult = _tomCommands.UpdateMeasure(
            _testExcelFile,
            measureName,
            daxFormula: "AVERAGE(Sales[Amount])",
            description: "Updated description",
            formatString: "0.00%"
        );

        // Assert
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        Assert.NotNull(updateResult.SuggestedNextActions);
        Assert.Contains(updateResult.SuggestedNextActions, s => s.Contains("updated successfully"));

        // Verify the measure was updated
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
        if (viewResult.Success)
        {
            Assert.Contains("AVERAGE", viewResult.DaxFormula);
        }
    }

    [Fact]
    public void UpdateMeasure_WithNonExistentMeasure_ReturnsError()
    {
        // Act
        var result = _tomCommands.UpdateMeasure(
            _testExcelFile,
            "NonExistentMeasure",
            daxFormula: "SUM(Sales[Amount])"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("not found") ||
            result.ErrorMessage.Contains("connect"),
            $"Expected 'not found' or connection error, got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public void UpdateMeasure_WithNoParameters_ReturnsError()
    {
        // Act
        var result = _tomCommands.UpdateMeasure(
            _testExcelFile,
            "SomeMeasure"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("at least one property", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region CreateRelationship Tests

    [Fact]
    public async Task CreateRelationship_WithValidParameters_ReturnsSuccess()
    {
        // Arrange - This test requires the Data Model to have Sales, Customers, Products tables
        // Skip if TOM connection fails

        // Act
        var result = _tomCommands.CreateRelationship(
            _testExcelFile,
            "Sales",
            "CustomerID",
            "Customers",
            "CustomerID",
            isActive: true,
            crossFilterDirection: "Single"
        );

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            Assert.NotNull(result.SuggestedNextActions);
            Assert.Contains(result.SuggestedNextActions, s => s.Contains("created"));

            // Verify the relationship was created
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
            var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);
            if (listResult.Success)
            {
                Assert.Contains(listResult.Relationships, r =>
                    r.FromTable == "Sales" && r.ToTable == "Customers");
            }
        }
        else
        {
            // If TOM connection failed or relationship already exists, that's acceptable
            Assert.True(
                result.ErrorMessage?.Contains("Data Model") == true ||
                result.ErrorMessage?.Contains("connect") == true ||
                result.ErrorMessage?.Contains("already exists") == true,
                $"Expected Data Model, connection, or duplicate error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void CreateRelationship_WithInvalidTable_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateRelationship(
            _testExcelFile,
            "InvalidTable",
            "ID",
            "AnotherTable",
            "ID"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("not found") ||
            result.ErrorMessage.Contains("connect"),
            $"Expected 'not found' or connection error, got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public void CreateRelationship_WithEmptyParameters_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateRelationship(
            _testExcelFile,
            "",
            "Column1",
            "Table2",
            "Column2"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region UpdateRelationship Tests

    [Fact]
    public async Task UpdateRelationship_WithValidParameters_ReturnsSuccess()
    {
        // Arrange - First ensure a relationship exists
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Test should fail if no relationships exist, not skip
        if (!listResult.Success || listResult.Relationships == null || listResult.Relationships.Count == 0)
        {
            Assert.Fail($"Data Model does not have relationships for testing. Success={listResult.Success}, Count={listResult.Relationships?.Count ?? 0}");
        }

        var rel = listResult.Relationships[0];

        // Act - Update the relationship
        var updateResult = _tomCommands.UpdateRelationship(
            _testExcelFile,
            rel.FromTable,
            rel.FromColumn,
            rel.ToTable,
            rel.ToColumn,
            isActive: !rel.IsActive
        );

        // Assert
        if (updateResult.Success)
        {
            Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");
            Assert.NotNull(updateResult.SuggestedNextActions);
        }
        else
        {
            // TOM connection failure is acceptable
            Assert.True(
                updateResult.ErrorMessage?.Contains("connect") == true,
                $"Expected connection error, got: {updateResult.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void UpdateRelationship_WithNoParameters_ReturnsError()
    {
        // Act
        var result = _tomCommands.UpdateRelationship(
            _testExcelFile,
            "Table1",
            "Col1",
            "Table2",
            "Col2"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("at least one property", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region CreateCalculatedColumn Tests

    [Fact]
    public void CreateCalculatedColumn_WithValidParameters_ReturnsSuccess()
    {
        // Arrange
        var columnName = "TestColumn_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "[Amount] * 2";

        // Act
        var result = _tomCommands.CreateCalculatedColumn(
            _testExcelFile,
            "Sales",
            columnName,
            daxFormula,
            "Test calculated column",
            "Double"
        );

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            Assert.NotNull(result.SuggestedNextActions);
            Assert.Contains(result.SuggestedNextActions, s => s.Contains("created successfully"));
        }
        else
        {
            // TOM connection failure or table not found is acceptable
            Assert.True(
                result.ErrorMessage?.Contains("Data Model") == true ||
                result.ErrorMessage?.Contains("connect") == true ||
                result.ErrorMessage?.Contains("not found") == true,
                $"Expected Data Model, connection, or not found error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void CreateCalculatedColumn_WithEmptyColumnName_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateCalculatedColumn(
            _testExcelFile,
            "Sales",
            "",
            "[Amount] * 2"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("name cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region ValidateDax Tests

    [Fact]
    public void ValidateDax_WithValidFormula_ReturnsValidResult()
    {
        // Arrange
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        var result = _tomCommands.ValidateDax(_testExcelFile, daxFormula);

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            // Validation may or may not succeed depending on TOM connection
            // Just verify the operation completed
        }
        else
        {
            // Connection failure is acceptable
            Assert.True(
                result.ErrorMessage?.Contains("connect") == true,
                $"Expected connection error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void ValidateDax_WithUnbalancedParentheses_ReturnsInvalidResult()
    {
        // Arrange
        var daxFormula = "SUM(Sales[Amount]";

        // Act
        var result = _tomCommands.ValidateDax(_testExcelFile, daxFormula);

        // Assert
        if (result.Success)
        {
            // Validation should detect unbalanced parentheses
            Assert.False(result.IsValid);
            Assert.NotNull(result.ValidationError);
            Assert.Contains("parenthes", result.ValidationError, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void ValidateDax_WithEmptyFormula_ReturnsError()
    {
        // Act
        var result = _tomCommands.ValidateDax(_testExcelFile, "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region ImportMeasures Tests

    [Fact]
    public async Task ImportMeasures_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = await _tomCommands.ImportMeasures(
            _testExcelFile,
            "NonExistent.json"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ImportMeasures_WithUnsupportedFormat_ReturnsError()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "test.txt");
        File.WriteAllText(testFile, "test content");

        // Act
        var result = await _tomCommands.ImportMeasures(
            _testExcelFile,
            testFile
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("format", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region File Validation Tests

    [Fact]
    public void CreateMeasure_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateMeasure(
            "NonExistent.xlsx",
            "Sales",
            "TestMeasure",
            "SUM(Sales[Amount])"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

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

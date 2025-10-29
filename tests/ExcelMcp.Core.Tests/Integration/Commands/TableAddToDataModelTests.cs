using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Dedicated tests for AddToDataModelAsync to catch edge cases and specific failure scenarios.
/// These tests aim to replicate real-world table configurations that might fail.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
public class TableAddToDataModelTests : IDisposable
{
    private readonly ITableCommands _tableCommands;
    private readonly IRangeCommands _rangeCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;
    private bool _disposed;

    public TableAddToDataModelTests(ITestOutputHelper output)
    {
        _output = output;
        _tableCommands = new TableCommands();
        _rangeCommands = new RangeCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_DataModel_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Test adding a simple table with basic data types (string, number, date).
    /// Data Model is always available in Excel 2013+, so this MUST succeed.
    /// </summary>
    [Fact]
    public async Task AddToDataModel_SimpleTable_MustSucceed()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "SimpleTable.xlsx");
        await CreateFileWithSimpleTable(testFile, "TestTable");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "TestTable");

        // Assert - MUST succeed, no excuses
        _output.WriteLine($"Result: Success={result.Success}, Error={result.ErrorMessage}");

        Assert.True(result.Success,
            $"AddToDataModel MUST succeed - Data Model is always available in Excel 2013+. " +
            $"This failure indicates the implementation is using the wrong API. " +
            $"Error: {result.ErrorMessage}");

        Assert.Equal("add-to-data-model", result.Action);
        Assert.Contains("added to Power Pivot Data Model", result.WorkflowHint ?? "");
    }

    /// <summary>
    /// Test adding a table with special characters in the name.
    /// Real-world scenario: User creates table named "Milestones" (from error report).
    /// Data Model is always available, so this MUST succeed.
    /// </summary>
    [Fact]
    public async Task AddToDataModel_TableWithCommonName_MustSucceed()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "Milestones.xlsx");
        await CreateFileWithSimpleTable(testFile, "Milestones");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "Milestones");

        // Assert - MUST succeed
        _output.WriteLine($"Result: Success={result.Success}, Error={result.ErrorMessage}");

        Assert.True(result.Success,
            $"AddToDataModel MUST succeed. Current implementation is broken (wrong API usage). " +
            $"Error: {result.ErrorMessage}");
        Assert.Equal("add-to-data-model", result.Action);
    }

    /// <summary>
    /// Test adding a table with numeric values only.
    /// Data Model is always available, so this MUST succeed.
    /// </summary>
    [Fact]
    public async Task AddToDataModel_TableWithNumericData_MustSucceed()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "NumericTable.xlsx");
        await CreateFileWithNumericTable(testFile, "Numbers");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "Numbers");

        // Assert - MUST succeed
        _output.WriteLine($"Result: Success={result.Success}, Error={result.ErrorMessage}");
        Assert.True(result.Success,
            $"AddToDataModel MUST succeed. Error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Test adding a table with empty cells.
    /// Data Model is always available, so this MUST succeed.
    /// </summary>
    [Fact]
    public async Task AddToDataModel_TableWithEmptyCells_MustSucceed()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "SparseTable.xlsx");
        await CreateFileWithSparseTable(testFile, "SparseData");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "SparseData");

        // Assert - MUST succeed
        _output.WriteLine($"Result: Success={result.Success}, Error={result.ErrorMessage}");
        Assert.True(result.Success,
            $"AddToDataModel MUST succeed. Error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Test adding a large table (100+ rows).
    /// Data Model is always available, so this MUST succeed.
    /// </summary>
    [Fact]
    public async Task AddToDataModel_LargeTable_MustSucceed()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "LargeTable.xlsx");
        await CreateFileWithLargeTable(testFile, "LargeData", rowCount: 100);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "LargeData");

        // Assert - MUST succeed
        _output.WriteLine($"Result: Success={result.Success}, Error={result.ErrorMessage}");
        Assert.True(result.Success,
            $"AddToDataModel MUST succeed. Error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Test adding a table with formula-based columns.
    /// Data Model is always available, so this MUST succeed.
    /// </summary>
    [Fact]
    public async Task AddToDataModel_TableWithFormulas_MustSucceed()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "FormulaTable.xlsx");
        await CreateFileWithFormulaTable(testFile, "Calculations");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "Calculations");

        // Assert - MUST succeed
        _output.WriteLine($"Result: Success={result.Success}, Error={result.ErrorMessage}");
        Assert.True(result.Success,
            $"AddToDataModel MUST succeed. Error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Test adding a table when the workbook already has other tables in Data Model.
    /// Data Model is always available, so this MUST succeed.
    /// </summary>
    [Fact]
    public async Task AddToDataModel_WhenOtherTablesExist_MustSucceed()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "MultipleTable.xlsx");
        await CreateFileWithMultipleTables(testFile);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - Try to add both tables
        var result1 = await _tableCommands.AddToDataModelAsync(batch, "Table1");
        _output.WriteLine($"Table1 Result: Success={result1.Success}, Error={result1.ErrorMessage}");

        var result2 = await _tableCommands.AddToDataModelAsync(batch, "Table2");
        _output.WriteLine($"Table2 Result: Success={result2.Success}, Error={result2.ErrorMessage}");

        // Assert - Both MUST succeed
        Assert.True(result1.Success,
            $"AddToDataModel MUST succeed for Table1. Error: {result1.ErrorMessage}");
        Assert.True(result2.Success,
            $"AddToDataModel MUST succeed for Table2. Error: {result2.ErrorMessage}");
    }

    /// <summary>
    /// Test adding the same table twice (should fail gracefully).
    /// </summary>
    [Fact]
    public async Task AddToDataModel_TableAlreadyInModel_ShouldFailGracefully()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "DuplicateAdd.xlsx");
        await CreateFileWithSimpleTable(testFile, "TestTable");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - First add
        var result1 = await _tableCommands.AddToDataModelAsync(batch, "TestTable");
        _output.WriteLine($"First add: Success={result1.Success}, Error={result1.ErrorMessage}");

        if (result1.Success)
        {
            // Act - Second add (should fail)
            var result2 = await _tableCommands.AddToDataModelAsync(batch, "TestTable");
            _output.WriteLine($"Second add: Success={result2.Success}, Error={result2.ErrorMessage}");

            // Assert
            Assert.False(result2.Success);
            Assert.Contains("already in the Data Model", result2.ErrorMessage ?? "");
        }
    }

    #region Helper Methods

    private async Task CreateFileWithSimpleTable(string filePath, string tableName)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create data
        var data = new List<List<object?>>
        {
            new() { "Name", "Value", "Date" },
            new() { "Item1", 100, DateTime.Now },
            new() { "Item2", 200, DateTime.Now.AddDays(-1) },
            new() { "Item3", 300, DateTime.Now.AddDays(-2) }
        };

        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:C4", data);

        // Create table
        var tableResult = await _tableCommands.CreateAsync(batch, "Sheet1", tableName, "A1:C4", true);
        Assert.True(tableResult.Success);

        await batch.SaveAsync();
    }

    private async Task CreateFileWithNumericTable(string filePath, string tableName)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create numeric-only data
        var data = new List<List<object?>>
        {
            new() { "Col1", "Col2", "Col3" },
            new() { 1, 10, 100 },
            new() { 2, 20, 200 },
            new() { 3, 30, 300 }
        };

        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:C4", data);

        var tableResult = await _tableCommands.CreateAsync(batch, "Sheet1", tableName, "A1:C4", true);
        Assert.True(tableResult.Success);

        await batch.SaveAsync();
    }

    private async Task CreateFileWithSparseTable(string filePath, string tableName)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create sparse data with nulls
        var data = new List<List<object?>>
        {
            new() { "Col1", "Col2", "Col3" },
            new() { "A", null, 100 },
            new() { null, "B", null },
            new() { "C", "D", 300 }
        };

        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:C4", data);

        var tableResult = await _tableCommands.CreateAsync(batch, "Sheet1", tableName, "A1:C4", true);
        Assert.True(tableResult.Success);

        await batch.SaveAsync();
    }

    private async Task CreateFileWithLargeTable(string filePath, string tableName, int rowCount)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create large dataset
        var data = new List<List<object?>>
        {
            new() { "ID", "Name", "Value" }
        };

        for (int i = 1; i <= rowCount; i++)
        {
            data.Add(new List<object?> { i, $"Item{i}", i * 10 });
        }

        var range = $"A1:C{rowCount + 1}";
        await _rangeCommands.SetValuesAsync(batch, "Sheet1", range, data);

        var tableResult = await _tableCommands.CreateAsync(batch, "Sheet1", tableName, range, true);
        Assert.True(tableResult.Success);

        await batch.SaveAsync();
    }

    private async Task CreateFileWithFormulaTable(string filePath, string tableName)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create data with base values
        var data = new List<List<object?>>
        {
            new() { "Base", "Multiplier", "Result" },
            new() { 10, 2, null },
            new() { 20, 3, null },
            new() { 30, 4, null }
        };

        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:C4", data);

        // Add formulas to Result column
        var formulas = new List<List<string>>
        {
            new() { "=A2*B2" },
            new() { "=A3*B3" },
            new() { "=A4*B4" }
        };
        await _rangeCommands.SetFormulasAsync(batch, "Sheet1", "C2:C4", formulas);

        var tableResult = await _tableCommands.CreateAsync(batch, "Sheet1", tableName, "A1:C4", true);
        Assert.True(tableResult.Success);

        await batch.SaveAsync();
    }

    private async Task CreateFileWithMultipleTables(string filePath)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create Table1
        var data1 = new List<List<object?>>
        {
            new() { "Name", "Value" },
            new() { "A", 1 },
            new() { "B", 2 }
        };
        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:B3", data1);
        var table1Result = await _tableCommands.CreateAsync(batch, "Sheet1", "Table1", "A1:B3", true);
        Assert.True(table1Result.Success);

        // Create Table2
        var data2 = new List<List<object?>>
        {
            new() { "Item", "Count" },
            new() { "X", 10 },
            new() { "Y", 20 }
        };
        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "D1:E3", data2);
        var table2Result = await _tableCommands.CreateAsync(batch, "Sheet1", "Table2", "D1:E3", true);
        Assert.True(table2Result.Success);

        await batch.SaveAsync();
    }

    private void AssertProperErrorOrSuccess(OperationResult result)
    {
        if (!result.Success)
        {
            var errorMsg = result.ErrorMessage ?? "";

            // CRITICAL: Generic COM errors MUST include context
            if (errorMsg.Contains("Value does not fall within the expected range"))
            {
                Assert.True(
                    errorMsg.Contains("Power Pivot") ||
                    errorMsg.Contains("Data Model") ||
                    errorMsg.Contains("connection") ||
                    errorMsg.Contains("Ensure"),
                    $"Generic COM error lacks context. Error: {errorMsg}");
            }

            // Acceptable failures
            bool isAcceptableFailure =
                errorMsg.Contains("Data Model not available") ||
                errorMsg.Contains("Power Pivot") ||
                errorMsg.Contains("does not have a Data Model") ||
                errorMsg.Contains("already in the Data Model") ||
                errorMsg.Contains("Connections.Add2") ||
                errorMsg.Contains("CreateModelConnection");

            Assert.True(isAcceptableFailure,
                $"Unexpected error type. Should be environment/configuration issue with clear guidance. Error: {errorMsg}");
        }
    }

    #endregion

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failure is non-critical
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}

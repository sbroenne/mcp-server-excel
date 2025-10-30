using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
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

        // Assert - MUST succeed
        Assert.True(result.Success, $"AddToDataModelAsync failed unexpectedly. Error: {result.ErrorMessage}");

        Assert.Equal("table-add-to-datamodel", result.Action);
        Assert.Contains("added to Power Pivot Data Model", result.WorkflowHint ?? "");
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
        var result2 = await _tableCommands.AddToDataModelAsync(batch, "Table2");

        // Assert - Both MUST succeed
        Assert.True(result1.Success, $"AddToDataModelAsync failed for Table1. Error: {result1.ErrorMessage}");
        Assert.True(result2.Success, $"AddToDataModelAsync failed for Table2. Error: {result2.ErrorMessage}");
    }

    /// <summary>
    /// Test adding the same table twice (should fail gracefully).
    /// </summary>
    [Fact]
    public async Task AddToDataModel_TableAlreadyInModel_ShouldFailGracefullyOnSecondAttempt()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "DuplicateAdd.xlsx");
        await CreateFileWithSimpleTable(testFile, "TestTable");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - First add
        var result1 = await _tableCommands.AddToDataModelAsync(batch, "TestTable");

        // Assert - First add MUST succeed
        Assert.True(result1.Success, $"First attempt to add table to model failed unexpectedly: {result1.ErrorMessage}");

        // Act - Second add (should fail)
        var result2 = await _tableCommands.AddToDataModelAsync(batch, "TestTable");

        // Assert - Second add MUST fail with a specific error
        Assert.False(result2.Success, "Adding the same table to the Data Model twice should fail, but it succeeded.");
        Assert.Contains("is already in the Data Model", result2.ErrorMessage ?? "");
    }

    #region Helper Methods

    private async Task CreateFileWithSimpleTable(string filePath, string tableName)
    {
        await CreateFileWithData(filePath, tableName,
        [
            new() { "Name", "Value", "Date" },
            new() { "Item1", 100, DateTime.Now },
            new() { "Item2", 200, DateTime.Now.AddDays(-1) },
            new() { "Item3", 300, DateTime.Now.AddDays(-2) }
        ], "A1:C4");
    }

    private async Task CreateFileWithLargeTable(string filePath, string tableName, int rowCount)
    {
        var data = new List<List<object?>> { new() { "ID", "Name", "Value" } };
        for (int i = 1; i <= rowCount; i++)
        {
            data.Add([i, $"Item{i}", i * 10]);
        }

        await CreateFileWithData(filePath, tableName, data, $"A1:C{rowCount + 1}");
    }

    private async Task CreateFileWithData(string filePath, string tableName, List<List<object?>> data, string range)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        await _rangeCommands.SetValuesAsync(batch, "Sheet1", range, data);

        var tableResult = await _tableCommands.CreateAsync(batch, "Sheet1", tableName, range, true);
        Assert.True(tableResult.Success);

        await batch.SaveAsync();
    }

    private async Task CreateFileWithMultipleTables(string filePath)
    {
        var createResult = await _fileCommands.CreateEmptyAsync(filePath);
        Assert.True(createResult.Success);

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        // Create Table1
        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:B3",
        [
            new() { "Name", "Value" },
            new() { "A", 1 },
            new() { "B", 2 }
        ]);
        var table1Result = await _tableCommands.CreateAsync(batch, "Sheet1", "Table1", "A1:B3", true);
        Assert.True(table1Result.Success);

        // Create Table2
        await _rangeCommands.SetValuesAsync(batch, "Sheet1", "D1:E3",
        [
            new() { "Item", "Count" },
            new() { "X", 10 },
            new() { "Y", 20 }
        ]);
        var table2Result = await _tableCommands.CreateAsync(batch, "Sheet1", "Table2", "D1:E3", true);
        Assert.True(table2Result.Success);

        await batch.SaveAsync();
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

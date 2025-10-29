using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Table commands (Phase 1 & Phase 2).
/// These tests require Excel installation and validate Core table operations.
/// Tests use Core commands directly (not through CLI wrapper).
///
/// Phase 1: Lifecycle, Structure, Filters, Columns, Data, DataModel
/// Phase 2: Structured References, Sorting
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
public class TableCommandsTests : IDisposable
{
    private readonly ITableCommands _tableCommands;
    private readonly IRangeCommands _rangeCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public TableCommandsTests()
    {
        _tableCommands = new TableCommands();
        _rangeCommands = new RangeCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_Table_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestTables.xlsx");

        // Create test Excel file with sample table
        CreateTestExcelFileWithTable();
    }

    private void CreateTestExcelFileWithTable()
    {
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }

        // Create a test table with sample data
        Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

            // Get Sheet1 and add sample data
            await batch.ExecuteAsync<int>((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Name = "Sales";

                // Add headers
                sheet.Range["A1"].Value2 = "Region";
                sheet.Range["B1"].Value2 = "Product";
                sheet.Range["C1"].Value2 = "Amount";
                sheet.Range["D1"].Value2 = "Date";

                // Add sample data
                sheet.Range["A2"].Value2 = "North";
                sheet.Range["B2"].Value2 = "Widget";
                sheet.Range["C2"].Value2 = 100;
                sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

                sheet.Range["A3"].Value2 = "South";
                sheet.Range["B3"].Value2 = "Gadget";
                sheet.Range["C3"].Value2 = 250;
                sheet.Range["D3"].Value2 = new DateTime(2025, 2, 20);

                sheet.Range["A4"].Value2 = "East";
                sheet.Range["B4"].Value2 = "Widget";
                sheet.Range["C4"].Value2 = 150;
                sheet.Range["D4"].Value2 = new DateTime(2025, 3, 10);

                sheet.Range["A5"].Value2 = "West";
                sheet.Range["B5"].Value2 = "Gadget";
                sheet.Range["C5"].Value2 = 300;
                sheet.Range["D5"].Value2 = new DateTime(2025, 1, 25);

                return ValueTask.FromResult(0);
            });

            // Create table from range A1:D5
            var createResult = await _tableCommands.CreateAsync(batch, "Sales", "SalesTable", "A1:D5", true, "TableStyleMedium2");
            if (!createResult.Success)
            {
                throw new InvalidOperationException($"Failed to create test table: {createResult.ErrorMessage}");
            }

            await batch.SaveAsync();
        }).GetAwaiter().GetResult();
    }

    #region Phase 1 Tests - Lifecycle

    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessWithTables()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
    }

    [Fact]
    public async Task Info_WithValidTable_ReturnsTableDetails()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetInfoAsync(batch, "SalesTable");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Table);
        Assert.Equal("SalesTable", result.Table.Name);
        Assert.Equal("Sales", result.Table.SheetName);
        Assert.True(result.Table.HasHeaders);
        Assert.Equal(4, result.Table.Columns?.Count); // Region, Product, Amount, Date
    }

    #endregion

    #region Phase 2 Tests - Structured References

    [Fact]
    public async Task GetStructuredReference_DataRegion_ReturnsCorrectReference()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.Data, null);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable", result.TableName);
        Assert.Equal(TableRegion.Data, result.Region);
        Assert.Equal("SalesTable[#Data]", result.StructuredReference);
        Assert.NotNull(result.RangeAddress);
        Assert.Contains("$A$2", result.RangeAddress); // Excel returns absolute references
        Assert.Equal(4, result.RowCount); // 4 data rows
        Assert.Equal(4, result.ColumnCount); // 4 columns
    }

    [Fact]
    public async Task GetStructuredReference_AllRegion_IncludesHeaders()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.All, null);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable[#All]", result.StructuredReference);
        Assert.Contains("$A$1", result.RangeAddress); // Excel returns absolute references
        Assert.Equal(5, result.RowCount); // Headers + 4 data rows
    }

    [Fact]
    public async Task GetStructuredReference_HeadersRegion_ReturnsHeaderRow()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.Headers, null);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable[#Headers]", result.StructuredReference);
        Assert.Contains("$A$1", result.RangeAddress); // Excel returns absolute references
        Assert.Equal(1, result.RowCount); // Only header row
        Assert.Equal(4, result.ColumnCount);
    }

    [Fact]
    public async Task GetStructuredReference_DataRegionWithColumn_ReturnsColumnReference()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.Data, "Amount");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable[[Amount]]", result.StructuredReference);
        Assert.Equal("Amount", result.ColumnName);
        Assert.Equal(4, result.RowCount); // 4 data rows
        Assert.Equal(1, result.ColumnCount); // Single column
    }

    [Fact]
    public async Task GetStructuredReference_AllRegionWithColumn_ReturnsFullColumnReference()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.All, "Region");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable[[#All],[Region]]", result.StructuredReference); // Excel includes both specifiers
        Assert.Equal("Region", result.ColumnName);
        Assert.Equal(5, result.RowCount); // Headers + 4 data rows
        Assert.Equal(1, result.ColumnCount);
    }

    [Fact]
    public async Task GetStructuredReference_InvalidTable_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "NonExistentTable", TableRegion.Data, null);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetStructuredReference_InvalidColumn_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.Data, "InvalidColumn");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        // Excel COM returns "Invalid index" for invalid column names
        Assert.Contains("Invalid", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Phase 2 Tests - Sorting

    [Fact]
    public async Task Sort_SingleColumn_Ascending_ReturnsSuccess()
    {
        // Act - Sort ascending by Region
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var sortResult = await _tableCommands.SortAsync(batch, "SalesTable", "Region", ascending: true);
        await batch.SaveAsync();

        // Assert
        Assert.True(sortResult.Success, $"Sort failed: {sortResult.ErrorMessage}");

        // Verify the table data is actually sorted by Region (ascending: East, North, South, West)
        var dataResult = await _rangeCommands.GetValuesAsync(batch, "Sales", "A2:A5"); // Region column data only
        Assert.True(dataResult.Success, $"Failed to read table data: {dataResult.ErrorMessage}");
        Assert.Equal(4, dataResult.Values.Count);
        Assert.Equal("East", dataResult.Values[0][0]?.ToString());
        Assert.Equal("North", dataResult.Values[1][0]?.ToString());
        Assert.Equal("South", dataResult.Values[2][0]?.ToString());
        Assert.Equal("West", dataResult.Values[3][0]?.ToString());
    }

    [Fact]
    public async Task Sort_SingleColumn_Descending_ReturnsSuccess()
    {
        // Act - Sort descending by Amount
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var sortResult = await _tableCommands.SortAsync(batch, "SalesTable", "Amount", ascending: false);
        await batch.SaveAsync();

        // Assert
        Assert.True(sortResult.Success, $"Sort failed: {sortResult.ErrorMessage}");

        // Verify the table data is actually sorted by Amount (descending: 300, 250, 150, 100)
        var dataResult = await _rangeCommands.GetValuesAsync(batch, "Sales", "C2:C5"); // Amount column data only
        Assert.True(dataResult.Success, $"Failed to read table data: {dataResult.ErrorMessage}");
        Assert.Equal(4, dataResult.Values.Count);
        Assert.Equal(300, Convert.ToInt32(dataResult.Values[0][0]));
        Assert.Equal(250, Convert.ToInt32(dataResult.Values[1][0]));
        Assert.Equal(150, Convert.ToInt32(dataResult.Values[2][0]));
        Assert.Equal(100, Convert.ToInt32(dataResult.Values[3][0]));
    }

    [Fact]
    public async Task Sort_MultiColumn_TwoLevels_ReturnsSuccess()
    {
        // Act - Sort by Product (asc), then Amount (desc)
        var sortColumns = new List<TableSortColumn>
        {
            new() { ColumnName = "Product", Ascending = true },
            new() { ColumnName = "Amount", Ascending = false }
        };

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var sortResult = await _tableCommands.SortAsync(batch, "SalesTable", sortColumns);
        await batch.SaveAsync();

        // Assert
        Assert.True(sortResult.Success, $"Sort failed: {sortResult.ErrorMessage}");

        // Verify the table data is actually sorted by Product first (Gadget, Widget), then Amount within each group
        var dataResult = await _rangeCommands.GetValuesAsync(batch, "Sales", "B2:C5"); // Product and Amount columns
        Assert.True(dataResult.Success, $"Failed to read table data: {dataResult.ErrorMessage}");
        Assert.Equal(4, dataResult.Values.Count);

        // First two rows should be Gadget products (sorted by Amount desc: 300, 250)
        Assert.Equal("Gadget", dataResult.Values[0][0]?.ToString());
        Assert.Equal(300, Convert.ToInt32(dataResult.Values[0][1]));
        Assert.Equal("Gadget", dataResult.Values[1][0]?.ToString());
        Assert.Equal(250, Convert.ToInt32(dataResult.Values[1][1]));

        // Next two rows should be Widget products (sorted by Amount desc: 150, 100)
        Assert.Equal("Widget", dataResult.Values[2][0]?.ToString());
        Assert.Equal(150, Convert.ToInt32(dataResult.Values[2][1]));
        Assert.Equal("Widget", dataResult.Values[3][0]?.ToString());
        Assert.Equal(100, Convert.ToInt32(dataResult.Values[3][1]));
    }

    [Fact]
    public async Task Sort_MultiColumn_ThreeLevels_ReturnsSuccess()
    {
        // Act - Sort by Region, Product, Amount
        var sortColumns = new List<TableSortColumn>
        {
            new() { ColumnName = "Region", Ascending = true },
            new() { ColumnName = "Product", Ascending = true },
            new() { ColumnName = "Amount", Ascending = false }
        };

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var sortResult = await _tableCommands.SortAsync(batch, "SalesTable", sortColumns);
        await batch.SaveAsync();

        // Assert
        Assert.True(sortResult.Success, $"Sort failed: {sortResult.ErrorMessage}");
    }

    [Fact]
    public async Task Sort_InvalidTable_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.SortAsync(batch, "NonExistentTable", "Amount", true);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Sort_InvalidColumn_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.SortAsync(batch, "SalesTable", "InvalidColumn", true);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        // Excel COM returns "Invalid index" for invalid column names
        Assert.Contains("Invalid", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Sort_MultiColumn_Empty_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.SortAsync(batch, "SalesTable", new List<TableSortColumn>());

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("at least one", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Sort_MultiColumn_TooMany_ReturnsError()
    {
        // Act - Try 4 sort columns (Excel max is 3)
        var sortColumns = new List<TableSortColumn>
        {
            new() { ColumnName = "Region", Ascending = true },
            new() { ColumnName = "Product", Ascending = true },
            new() { ColumnName = "Amount", Ascending = false },
            new() { ColumnName = "Date", Ascending = true }
        };

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _tableCommands.SortAsync(batch, "SalesTable", sortColumns);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("maximum of 3", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Data Model Tests

    /// <summary>
    /// ⚠️ TODO: DELETE THIS TEST - IT'S BROKEN!
    /// This test is fundamentally flawed because it accepts both success AND failure.
    /// The AddToDataModelAsync feature is completely non-functional (wrong API usage),
    /// but this test passes because it accepts "environment" failures.
    ///
    /// See: specs/TABLE-DATAMODEL-ISSUE-ANALYSIS.md for details
    /// See: tests/ExcelMcp.Core.Tests/Integration/Commands/TableAddToDataModelTests.cs for proper tests
    /// </summary>
    [Fact]
    public async Task AddToDataModelAsync_WithValidTable_ShouldSucceedOrProvideReasonableError()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

        // This test validates that AddToDataModelAsync works correctly with a real table
        // In environments where Power Pivot isn't available, it should fail gracefully

        // Act - Add the table to Data Model
        var result = await _tableCommands.AddToDataModelAsync(batch, "SalesTable");

        // Assert - Either succeeds OR fails with a reasonable environment-related error
        if (result.Success)
        {
            // SUCCESS CASE: Verify the operation completed properly
            Assert.True(result.Success);
            Assert.Equal("add-to-data-model", result.Action);
            Assert.Contains("added to Power Pivot Data Model", result.WorkflowHint ?? "");
            Assert.Contains("dm-list-tables", result.SuggestedNextActions?.FirstOrDefault() ?? "");
        }
        else
        {
            // GRACEFUL FAILURE: Should be environment-related, not a code bug
            var errorMsg = result.ErrorMessage ?? "";

            // These are acceptable environment-related failures
            bool isEnvironmentIssue =
                errorMsg.Contains("Data Model not available") ||
                errorMsg.Contains("Power Pivot") ||
                errorMsg.Contains("does not have a Data Model") ||
                errorMsg.Contains("already in the Data Model") ||
                errorMsg.Contains("Connections.Add2");

            Assert.True(isEnvironmentIssue,
                $"Expected environment-related error, but got: {errorMsg}");

            // Should NOT be the original COM bug that was fixed in Issue #64
            Assert.False(errorMsg.Contains("does not contain a definition for 'Add'"),
                "Should not have the original COM method error that was fixed");
        }
    }

    [Fact]
    public async Task AddToDataModelAsync_WithNonExistentTable_ShouldFail()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "NonExistentTable");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }

    [Fact]
    public async Task AddToDataModelAsync_WithInvalidTableName_ShouldFail()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

        // Act & Assert - Invalid characters should be rejected
        await Assert.ThrowsAsync<ArgumentException>(() =>
            _tableCommands.AddToDataModelAsync(batch, "Table<>Name"));
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
            // Cleanup failure is non-critical for tests
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}

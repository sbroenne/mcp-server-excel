using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Tests for advanced Table operations (totals, filters, data operations, columns)
/// Optimized: Single batch per test, no SaveAsync unless testing persistence
/// </summary>
public partial class TableCommandsTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task ToggleTotals_EnableTotals_AddsTotalsRow()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(ToggleTotals_EnableTotals_AddsTotalsRow));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.ToggleTotalsAsync(batch, "TestTable", true);

        // Assert
        Assert.True(result.Success, $"ToggleTotals failed: {result.ErrorMessage}");
        
        // Verify totals row added
        var info = await _tableCommands.GetInfoAsync(batch, "TestTable");
        Assert.True(info.Success);
        Assert.True(info.Table.ShowTotals);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetColumnTotal_WithSumFunction_SetsTotalFormula()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(SetColumnTotal_WithSumFunction_SetsTotalFormula));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Enable totals first
        await _tableCommands.ToggleTotalsAsync(batch, "TestTable", true);

        // Act - Set sum for Amount column
        var result = await _tableCommands.SetColumnTotalAsync(batch, "TestTable", "Amount", "Sum");

        // Assert
        Assert.True(result.Success, $"SetColumnTotal failed: {result.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task AppendRows_WithNewData_AddsRowsToTable()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(AppendRows_WithNewData_AddsRowsToTable));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        var newRows = new List<List<object?>>
        {
            new() { "West", "Widget", 500, DateTime.Now },
            new() { "East", "Gadget", 600, DateTime.Now }
        };

        // Act
        var result = await _tableCommands.AppendRowsAsync(batch, "TestTable", newRows);

        // Assert
        Assert.True(result.Success, $"AppendRows failed: {result.ErrorMessage}");
        
        // Verify table grew (AppendRowsAsync returns OperationResult, not row count)
        var info = await _tableCommands.GetInfoAsync(batch, "TestTable");
        Assert.True(info.Success);
        Assert.True(info.Table.RowCount >= 4); // Original 2 + appended 2
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task ApplyFilter_WithColumnCriteria_FiltersTable()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(ApplyFilter_WithColumnCriteria_FiltersTable));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - Filter Region column to show only "North"
        var result = await _tableCommands.ApplyFilterAsync(batch, "TestTable", "Region", new List<string> { "North" });

        // Assert
        Assert.True(result.Success, $"ApplyFilter failed: {result.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task ClearFilters_AfterFiltering_RemovesAllFilters()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(ClearFilters_AfterFiltering_RemovesAllFilters));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Apply a filter first
        await _tableCommands.ApplyFilterAsync(batch, "TestTable", "Region", new List<string> { "North" });

        // Act - Clear all filters
        var result = await _tableCommands.ClearFiltersAsync(batch, "TestTable");

        // Assert
        Assert.True(result.Success, $"ClearFilters failed: {result.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetFilters_WithActiveFilters_ReturnsFilterInfo()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(GetFilters_WithActiveFilters_ReturnsFilterInfo));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Apply a filter
        await _tableCommands.ApplyFilterAsync(batch, "TestTable", "Region", new List<string> { "North" });

        // Act
        var result = await _tableCommands.GetFiltersAsync(batch, "TestTable");

        // Assert
        Assert.True(result.Success, $"GetFilters failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.ColumnFilters);
        Assert.Contains(result.ColumnFilters, f => f.ColumnName == "Region");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task RemoveColumn_ExistingColumn_DeletesColumn()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(RemoveColumn_ExistingColumn_DeletesColumn));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - Remove Date column
        var result = await _tableCommands.RemoveColumnAsync(batch, "TestTable", "Date");

        // Assert
        Assert.True(result.Success, $"RemoveColumn failed: {result.ErrorMessage}");
        
        // Verify column removed
        var info = await _tableCommands.GetInfoAsync(batch, "TestTable");
        Assert.True(info.Success);
        Assert.DoesNotContain("Date", info.Table.Columns);
        Assert.Equal(3, info.Table.Columns?.Count); // Should have 3 columns now
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task RenameColumn_ExistingColumn_ChangesColumnName()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(RenameColumn_ExistingColumn_ChangesColumnName));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - Rename Amount to Sales
        var result = await _tableCommands.RenameColumnAsync(batch, "TestTable", "Amount", "Sales");

        // Assert
        Assert.True(result.Success, $"RenameColumn failed: {result.ErrorMessage}");
        
        // Verify rename
        var info = await _tableCommands.GetInfoAsync(batch, "TestTable");
        Assert.True(info.Success);
        Assert.Contains("Sales", info.Table.Columns);
        Assert.DoesNotContain("Amount", info.Table.Columns);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Sort_ByColumn_SortsTableData()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(Sort_ByColumn_SortsTableData));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - Sort by Region ascending
        var result = await _tableCommands.SortAsync(batch, "TestTable", "Region", true);

        // Assert
        Assert.True(result.Success, $"Sort failed: {result.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetColumnNumberFormat_ExistingColumn_ReturnsFormat()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(GetColumnNumberFormat_ExistingColumn_ReturnsFormat));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.GetColumnNumberFormatAsync(batch, "TestTable", "Amount");

        // Assert
        Assert.True(result.Success, $"GetColumnNumberFormat failed: {result.ErrorMessage}");
        Assert.NotNull(result.Formats);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetColumnNumberFormat_WithCurrencyFormat_AppliesFormat()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(SetColumnNumberFormat_WithCurrencyFormat_AppliesFormat));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act - Set currency format for Amount column
        var result = await _tableCommands.SetColumnNumberFormatAsync(batch, "TestTable", "Amount", "$#,##0.00");

        // Assert
        Assert.True(result.Success, $"SetColumnNumberFormat failed: {result.ErrorMessage}");
        
        // Verify format applied
        var getResult = await _tableCommands.GetColumnNumberFormatAsync(batch, "TestTable", "Amount");
        Assert.True(getResult.Success);
        Assert.NotNull(getResult.Formats);
        // Format should be applied to all cells in the column
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task AddToDataModel_ExistingTable_AddsTableToModel()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(AddToDataModel_ExistingTable_AddsTableToModel));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _tableCommands.AddToDataModelAsync(batch, "TestTable");

        // Assert
        Assert.True(result.Success, $"AddToDataModel failed: {result.ErrorMessage}");
    }
}

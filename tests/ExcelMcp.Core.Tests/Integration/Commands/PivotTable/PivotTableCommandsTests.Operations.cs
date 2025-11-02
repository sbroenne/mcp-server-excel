using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable operations (List, GetInfo, Delete, Refresh, GetData)
/// Optimized: Single batch per test, no SaveAsync() unless testing persistence
/// </summary>
public partial class PivotTableCommandsTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(List_EmptyWorkbook_ReturnsEmptyList));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.PivotTables);
        Assert.Empty(result.PivotTables);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task List_WithPivotTable_ReturnsPivotTableInfo()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(List_WithPivotTable_ReturnsPivotTableInfo));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - No save needed, same batch
        var result = await _pivotCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.PivotTables);
        var pivot = Assert.Single(result.PivotTables);
        Assert.Equal("TestPivot", pivot.Name);
        Assert.Equal("SalesData", pivot.SheetName);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetInfo_ExistingPivotTable_ReturnsCompleteInfo()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(GetInfo_ExistingPivotTable_ReturnsCompleteInfo));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - No save needed
        var result = await _pivotCommands.GetInfoAsync(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"GetInfo failed: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTable.Name);
        Assert.NotEmpty(result.Fields);
        Assert.Equal(4, result.Fields.Count); // Region, Product, Sales, Date
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetInfo_NonExistentPivotTable_ReturnsError()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(GetInfo_NonExistentPivotTable_ReturnsError));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.GetInfoAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Delete_ExistingPivotTable_RemovesPivotTable()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(Delete_ExistingPivotTable_RemovesPivotTable));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - Delete in same batch
        var deleteResult = await _pivotCommands.DeleteAsync(batch, "TestPivot");
        
        // Assert
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");
        
        // Verify pivot no longer exists
        var listResult = await _pivotCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Empty(listResult.PivotTables);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Delete_NonExistentPivotTable_ReturnsError()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(Delete_NonExistentPivotTable_ReturnsError));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.DeleteAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Refresh_ExistingPivotTable_UpdatesData()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(Refresh_ExistingPivotTable_UpdatesData));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - Refresh in same batch
        var result = await _pivotCommands.RefreshAsync(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.True(result.RefreshTime <= DateTime.Now);
        Assert.True(result.SourceRecordCount >= 0);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetData_ExistingPivotTable_ReturnsData()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(GetData_ExistingPivotTable_ReturnsData));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create pivot with row field to generate data
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);
        
        // Add Region to row area
        var addRowResult = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");
        Assert.True(addRowResult.Success);

        // Act - GetData in same batch
        var result = await _pivotCommands.GetDataAsync(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"GetData failed: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.NotNull(result.Values);
        Assert.NotEmpty(result.Values);
    }
}

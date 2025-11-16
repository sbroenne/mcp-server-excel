using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable operations (List, GetInfo, Delete, Refresh, GetData)
/// Optimized: Single batch per test, no SaveAsync() unless testing persistence
/// </summary>
public partial class PivotTableCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(List_EmptyWorkbook_ReturnsEmptyList));

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _pivotCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.PivotTables);
        Assert.Empty(result.PivotTables);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task List_WithPivotTable_ReturnsPivotTableInfo()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(List_WithPivotTable_ReturnsPivotTableInfo));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - No save needed, same batch
        var result = _pivotCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.PivotTables);
        var pivot = Assert.Single(result.PivotTables);
        Assert.Equal("TestPivot", pivot.Name);
        Assert.Equal("SalesData", pivot.SheetName);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetInfo_ExistingPivotTable_ReturnsCompleteInfo()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(GetInfo_ExistingPivotTable_ReturnsCompleteInfo));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - No save needed
        var result = _pivotCommands.Read(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"GetInfo failed: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTable.Name);
        Assert.NotEmpty(result.Fields);
        Assert.Equal(4, result.Fields.Count); // Region, Product, Sales, Date
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetInfo_NonExistentPivotTable_ReturnsError()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(GetInfo_NonExistentPivotTable_ReturnsError));

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _pivotCommands.Read(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Delete_ExistingPivotTable_RemovesPivotTable()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(Delete_ExistingPivotTable_RemovesPivotTable));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - Delete in same batch
        var deleteResult = _pivotCommands.Delete(batch, "TestPivot");

        // Assert
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

        // Verify pivot no longer exists
        var listResult = _pivotCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Empty(listResult.PivotTables);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Delete_NonExistentPivotTable_ReturnsError()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(Delete_NonExistentPivotTable_ReturnsError));

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _pivotCommands.Delete(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Refresh_ExistingPivotTable_UpdatesData()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(Refresh_ExistingPivotTable_UpdatesData));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - Refresh in same batch
        var result = _pivotCommands.Refresh(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.True(result.RefreshTime <= DateTime.Now);
        Assert.True(result.SourceRecordCount >= 0);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task GetData_ExistingPivotTable_ReturnsData()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(GetData_ExistingPivotTable_ReturnsData));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create pivot with row field to generate data
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Region to row area
        var addRowResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        Assert.True(addRowResult.Success);

        // Act - GetData in same batch
        var result = _pivotCommands.GetData(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"GetData failed: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.NotNull(result.Values);
        Assert.NotEmpty(result.Values);
    }
}

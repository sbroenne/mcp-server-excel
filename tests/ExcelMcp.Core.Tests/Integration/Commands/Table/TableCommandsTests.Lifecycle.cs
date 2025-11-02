using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Tests for Table lifecycle operations (list, create, info)
/// </summary>
public partial class TableCommandsTests
{
    [Fact]
    [Trait("Speed", "Fast")]
    public async Task List_WithValidFile_ReturnsSuccessWithTables()
    {
        // Act - Use shared fixture file
        await using var batch = await ExcelSession.BeginBatchAsync(_tableFile);
        var result = await _tableCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Info_WithValidTable_ReturnsTableDetails()
    {
        // Act - Use shared fixture file
        await using var batch = await ExcelSession.BeginBatchAsync(_tableFile);
        var result = await _tableCommands.GetInfoAsync(batch, "SalesTable");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Table);
        Assert.Equal("SalesTable", result.Table.Name);
        Assert.Equal("Sales", result.Table.SheetName);
        Assert.True(result.Table.HasHeaders);
        Assert.Equal(4, result.Table.Columns?.Count); // Region, Product, Amount, Date
    }

    [Fact]
    public async Task Create_WithValidData_CreatesTable()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(TableCommandsTests), nameof(Create_WithValidData_CreatesTable), _tempDir);

        // Act - Use single batch for create data, create table, and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Add data first
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Name";
            sheet.Range["B1"].Value2 = "Value";
            sheet.Range["A2"].Value2 = "Test1";
            sheet.Range["B2"].Value2 = 100;
            sheet.Range["A3"].Value2 = "Test2";
            sheet.Range["B3"].Value2 = 200;
            return 0;
        });

        // Create table
        var result = await _tableCommands.CreateAsync(batch, "Sheet1", "TestTable", "A1:B3", true, "TableStyleLight1");
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");

        // Verify table was created
        var listResult = await _tableCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Tables, t => t.Name == "TestTable");
    }

    [Fact]
    public async Task Delete_WithExistingTable_RemovesTable()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(Delete_WithExistingTable_RemovesTable));

        // Act - Use single batch for delete and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Delete the table
        var result = await _tableCommands.DeleteAsync(batch, "SalesTable");
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify table was deleted
        var listResult = await _tableCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
    }

    [Fact]
    public async Task Rename_WithExistingTable_RenamesSuccessfully()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(Rename_WithExistingTable_RenamesSuccessfully));

        // Act - Use single batch for rename and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Rename the table
        var result = await _tableCommands.RenameAsync(batch, "SalesTable", "RevenueTable");
        Assert.True(result.Success, $"Rename failed: {result.ErrorMessage}");

        // Verify table was renamed
        var listResult = await _tableCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
        Assert.Contains(listResult.Tables, t => t.Name == "RevenueTable");
    }

    [Fact]
    public async Task Resize_WithExistingTable_ResizesSuccessfully()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(Resize_WithExistingTable_ResizesSuccessfully));

        // Act - Use single batch for resize and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Get initial size
        var initialInfo = await _tableCommands.GetInfoAsync(batch, "SalesTable");
        Assert.True(initialInfo.Success);
        var initialRowCount = initialInfo.Table!.RowCount;

        // Resize to A1:D10 (expand)
        var result = await _tableCommands.ResizeAsync(batch, "SalesTable", "A1:D10");
        Assert.True(result.Success, $"Resize failed: {result.ErrorMessage}");

        // Verify new size
        var resizedInfo = await _tableCommands.GetInfoAsync(batch, "SalesTable");
        Assert.True(resizedInfo.Success);
        Assert.Equal(9, resizedInfo.Table!.RowCount); // 10 rows - 1 header
    }

    [Fact]
    public async Task SetStyle_WithExistingTable_ChangesStyleSuccessfully()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(SetStyle_WithExistingTable_ChangesStyleSuccessfully));

        // Act - Use single batch for style change and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Change table style
        var result = await _tableCommands.SetStyleAsync(batch, "SalesTable", "TableStyleMedium2");
        Assert.True(result.Success, $"SetStyle failed: {result.ErrorMessage}");

        // Verify style was changed
        var info = await _tableCommands.GetInfoAsync(batch, "SalesTable");
        Assert.True(info.Success);
        Assert.Equal("TableStyleMedium2", info.Table!.TableStyle);
    }

    [Fact]
    public async Task AddColumn_WithExistingTable_AddsColumnSuccessfully()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(AddColumn_WithExistingTable_AddsColumnSuccessfully));

        // Act - Use single batch for add column and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Get initial column count
        var initialInfo = await _tableCommands.GetInfoAsync(batch, "SalesTable");
        Assert.True(initialInfo.Success);
        var initialColumnCount = initialInfo.Table!.Columns!.Count;

        // Add new column
        var result = await _tableCommands.AddColumnAsync(batch, "SalesTable", "NewColumn");
        Assert.True(result.Success, $"AddColumn failed: {result.ErrorMessage}");

        // Verify column was added
        var updatedInfo = await _tableCommands.GetInfoAsync(batch, "SalesTable");
        Assert.True(updatedInfo.Success);
        Assert.Equal(initialColumnCount + 1, updatedInfo.Table!.Columns!.Count);
        Assert.Contains("NewColumn", updatedInfo.Table.Columns);
    }
}

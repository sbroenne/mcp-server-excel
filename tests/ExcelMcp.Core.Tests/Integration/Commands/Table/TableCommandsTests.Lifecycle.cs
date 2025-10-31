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
    public async Task List_WithValidFile_ReturnsSuccessWithTables()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(List_WithValidFile_ReturnsSuccessWithTables));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
    }

    [Fact]
    public async Task Info_WithValidTable_ReturnsTableDetails()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(Info_WithValidTable_ReturnsTableDetails));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
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

        await batch.SaveAsync();
    }
}

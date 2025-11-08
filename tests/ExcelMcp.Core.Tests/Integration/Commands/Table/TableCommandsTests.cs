using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Integration tests for Table operations focusing on LLM use cases.
/// Tests cover essential workflows: create, list, info, delete, rename, resize, columns, filters, totals.
/// Uses TableTestsFixture which creates ONE Table file per test class.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Medium")]
public class TableCommandsTests : IClassFixture<TableTestsFixture>
{
    private readonly TableCommands _tableCommands;
    private readonly IRangeCommands _rangeCommands;
    private readonly string _tableFile;
    private readonly TableCreationResult _creationResult;
    private readonly string _tempDir;

    /// <summary>
    /// Initializes a new instance of the <see cref="TableCommandsTests"/> class.
    /// </summary>
    public TableCommandsTests(TableTestsFixture fixture)
    {
        _tableCommands = new TableCommands();
        _rangeCommands = new RangeCommands();
        _tableFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
        _tempDir = Path.GetDirectoryName(fixture.TestFilePath)!;
    }

    #region Core Lifecycle Tests (7 tests)

    /// <summary>
    /// Validates that the fixture creation succeeded.
    /// LLM use case: "create a table from a range"
    /// </summary>
    [Fact]
    public void Create_ViaFixture_CreatesTable()
    {
        Assert.True(_creationResult.Success,
            $"Table creation failed during fixture initialization: {_creationResult.ErrorMessage}");
        Assert.True(_creationResult.FileCreated);
        Assert.Equal(1, _creationResult.TablesCreated);
    }

    /// <summary>
    /// Tests listing tables in a workbook.
    /// LLM use case: "show me all tables in this workbook"
    /// </summary>
    [Fact]
    public async Task List_WithValidFile_ReturnsTables()
    {
        await using var batch = await ExcelSession.BeginBatchAsync(_tableFile);
        var result = await _tableCommands.ListAsync(batch);

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
    }

    /// <summary>
    /// Tests getting table details.
    /// LLM use case: "show me information about this table"
    /// </summary>
    [Fact]
    public async Task Info_WithValidTable_ReturnsTableDetails()
    {
        await using var batch = await ExcelSession.BeginBatchAsync(_tableFile);
        var result = await _tableCommands.GetAsync(batch, "SalesTable");

        Assert.True(result.Success);
        Assert.NotNull(result.Table);
        Assert.Equal("SalesTable", result.Table.Name);
        Assert.Equal("Sales", result.Table.SheetName);
        Assert.True(result.Table.HasHeaders);
        Assert.Equal(4, result.Table.Columns?.Count);
    }

    /// <summary>
    /// Tests creating a new table.
    /// LLM use case: "convert this range to a table"
    /// </summary>
    [Fact]
    public async Task Create_WithValidData_CreatesTable()
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(TableCommandsTests), nameof(Create_WithValidData_CreatesTable), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Add data first
        await batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = "Name";
            sheet.Range["B1"].Value2 = "Value";
            sheet.Range["A2"].Value2 = "Test1";
            sheet.Range["B2"].Value2 = 100;
            return 0;
        });

        // Create table
        var result = await _tableCommands.CreateAsync(batch, "Sheet1", "TestTable", "A1:B2", true, "TableStyleLight1");
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");

        // Verify table was created
        var listResult = await _tableCommands.ListAsync(batch);
        Assert.Contains(listResult.Tables, t => t.Name == "TestTable");
    }

    /// <summary>
    /// Tests deleting a table.
    /// LLM use case: "delete this table"
    /// </summary>
    [Fact]
    public async Task Delete_WithExistingTable_RemovesTable()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(Delete_WithExistingTable_RemovesTable));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.DeleteAsync(batch, "SalesTable");
        Assert.True(result.Success);

        // Verify deletion
        var listResult = await _tableCommands.ListAsync(batch);
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
    }

    /// <summary>
    /// Tests renaming a table.
    /// LLM use case: "rename this table"
    /// </summary>
    [Fact]
    public async Task Rename_WithExistingTable_RenamesSuccessfully()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(Rename_WithExistingTable_RenamesSuccessfully));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.RenameAsync(batch, "SalesTable", "RevenueTable");
        Assert.True(result.Success);

        // Verify rename
        var listResult = await _tableCommands.ListAsync(batch);
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
        Assert.Contains(listResult.Tables, t => t.Name == "RevenueTable");
    }

    /// <summary>
    /// Tests resizing a table.
    /// LLM use case: "expand this table to include more rows"
    /// </summary>
    [Fact]
    public async Task Resize_WithExistingTable_ResizesSuccessfully()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(Resize_WithExistingTable_ResizesSuccessfully));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var initialInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.True(initialInfo.Success);

        var result = await _tableCommands.ResizeAsync(batch, "SalesTable", "A1:D10");
        Assert.True(result.Success);

        // Verify resize
        var resizedInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.Equal(9, resizedInfo.Table!.RowCount); // 10 rows - 1 header
    }

    #endregion

    #region Column Operations (2 tests)

    /// <summary>
    /// Tests adding a column to a table.
    /// LLM use case: "add a new column to this table"
    /// </summary>
    [Fact]
    public async Task AddColumn_WithExistingTable_AddsColumnSuccessfully()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(AddColumn_WithExistingTable_AddsColumnSuccessfully));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var initialInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        var initialColumnCount = initialInfo.Table!.Columns!.Count;

        var result = await _tableCommands.AddColumnAsync(batch, "SalesTable", "NewColumn");
        Assert.True(result.Success);

        // Verify column added
        var updatedInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.Equal(initialColumnCount + 1, updatedInfo.Table!.Columns!.Count);
        Assert.Contains("NewColumn", updatedInfo.Table.Columns);
    }

    /// <summary>
    /// Tests renaming a column in a table.
    /// LLM use case: "rename this table column"
    /// </summary>
    [Fact]
    public async Task RenameColumn_WithExistingColumn_RenamesSuccessfully()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(RenameColumn_WithExistingColumn_RenamesSuccessfully));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var result = await _tableCommands.RenameColumnAsync(batch, "SalesTable", "Amount", "Revenue");
        Assert.True(result.Success);

        // Verify rename
        var info = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.Contains("Revenue", info.Table!.Columns!);
        Assert.DoesNotContain("Amount", info.Table.Columns);
    }

    #endregion

    #region Data Operations (2 tests)

    /// <summary>
    /// Tests appending rows to a table.
    /// LLM use case: "add these rows to the table"
    /// </summary>
    [Fact]
    public async Task Append_WithNewData_AddsRowsToTable()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(Append_WithNewData_AddsRowsToTable));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var newRows = new List<List<object?>>
        {
            new() { "West", "Widget", 500, DateTime.Now },
            new() { "East", "Gadget", 600, DateTime.Now }
        };

        var result = await _tableCommands.AppendAsync(batch, "SalesTable", newRows);
        Assert.True(result.Success);

        // Verify rows added
        var info = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.True(info.Table!.RowCount >= 6); // Original 4 + appended 2
    }

    /// <summary>
    /// Tests getting structured reference for a table column.
    /// LLM use case: "get the structured reference formula for this table column"
    /// </summary>
    [Fact]
    public async Task GetStructuredReference_WithValidTable_ReturnsReference()
    {
        await using var batch = await ExcelSession.BeginBatchAsync(_tableFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.Data, "Amount");

        Assert.True(result.Success);
        Assert.NotNull(result.StructuredReference);
        Assert.Contains("SalesTable", result.StructuredReference);
        Assert.Contains("Amount", result.StructuredReference);
    }

    #endregion

    #region Filter Operations (2 tests)

    /// <summary>
    /// Tests applying a filter to a table column.
    /// LLM use case: "filter this table to show only these values"
    /// </summary>
    [Fact]
    public async Task ApplyFilter_WithColumnCriteria_FiltersTable()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(ApplyFilter_WithColumnCriteria_FiltersTable));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.ApplyFilterAsync(batch, "SalesTable", "Region", ["North"]);

        Assert.True(result.Success);
    }

    /// <summary>
    /// Tests clearing all filters from a table.
    /// LLM use case: "remove all filters from this table"
    /// </summary>
    [Fact]
    public async Task ClearFilters_AfterFiltering_RemovesAllFilters()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(ClearFilters_AfterFiltering_RemovesAllFilters));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Apply filter first
        await _tableCommands.ApplyFilterAsync(batch, "SalesTable", "Region", ["North"]);

        // Clear filters
        var result = await _tableCommands.ClearFiltersAsync(batch, "SalesTable");
        Assert.True(result.Success);
    }

    #endregion

    #region Totals Operations (2 tests)

    /// <summary>
    /// Tests enabling totals row on a table.
    /// LLM use case: "add a totals row to this table"
    /// </summary>
    [Fact]
    public async Task ToggleTotals_EnableTotals_AddsTotalsRow()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(ToggleTotals_EnableTotals_AddsTotalsRow));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.ToggleTotalsAsync(batch, "SalesTable", true);

        Assert.True(result.Success);

        // Verify totals enabled
        var info = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.True(info.Table!.ShowTotals);
    }

    /// <summary>
    /// Tests setting a total function on a column.
    /// LLM use case: "set the total for this column to sum"
    /// </summary>
    [Fact]
    public async Task SetColumnTotal_WithSumFunction_SetsTotalFormula()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(SetColumnTotal_WithSumFunction_SetsTotalFormula));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Enable totals first
        await _tableCommands.ToggleTotalsAsync(batch, "SalesTable", true);

        // Set sum for Amount column
        var result = await _tableCommands.SetColumnTotalAsync(batch, "SalesTable", "Amount", "Sum");
        Assert.True(result.Success);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Creates a unique test file with SalesTable for modification tests.
    /// </summary>
    private async Task<string> CreateTestFileWithTableAsync(string testName)
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(TableCommandsTests), testName, _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create worksheet with sample data
        await batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Name = "Sales";

            // Headers
            sheet.Range["A1"].Value2 = "Region";
            sheet.Range["B1"].Value2 = "Product";
            sheet.Range["C1"].Value2 = "Amount";
            sheet.Range["D1"].Value2 = "Date";

            // Sample data
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

            return 0;
        });

        // Create table from range A1:D5
        var createResult = await _tableCommands.CreateAsync(batch, "Sales", "SalesTable", "A1:D5", true, "TableStyleMedium2");
        if (!createResult.Success)
        {
            throw new InvalidOperationException($"Failed to create test table: {createResult.ErrorMessage}");
        }

        await batch.SaveAsync();
        return testFile;
    }

    #endregion

    #region Numeric Column Name Tests (3 tests)

    /// <summary>
    /// Tests adding a column with a purely numeric name.
    /// LLM use case: "add a column named 60 for 60 months data"
    /// Regression test for: Column names can be numeric (e.g. 60 for 60 months)
    /// </summary>
    [Fact]
    public async Task AddColumn_WithNumericName_AddsColumnSuccessfully()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(AddColumn_WithNumericName_AddsColumnSuccessfully));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var initialInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        var initialColumnCount = initialInfo.Table!.Columns!.Count;

        // Add column with purely numeric name
        var result = await _tableCommands.AddColumnAsync(batch, "SalesTable", "60");
        Assert.True(result.Success, $"Failed to add numeric column: {result.ErrorMessage}");

        // Verify column added
        var updatedInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.Equal(initialColumnCount + 1, updatedInfo.Table!.Columns!.Count);
        Assert.Contains("60", updatedInfo.Table.Columns);
    }

    /// <summary>
    /// Tests renaming a column to a purely numeric name.
    /// LLM use case: "rename this column to 12 for 12 months"
    /// Regression test for: Column names can be numeric (e.g. 60 for 60 months)
    /// </summary>
    [Fact]
    public async Task RenameColumn_ToNumericName_RenamesSuccessfully()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(RenameColumn_ToNumericName_RenamesSuccessfully));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Rename "Amount" column to numeric name "60"
        var result = await _tableCommands.RenameColumnAsync(batch, "SalesTable", "Amount", "60");
        Assert.True(result.Success, $"Failed to rename to numeric column name: {result.ErrorMessage}");

        // Verify column renamed
        var updatedInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.Contains("60", updatedInfo.Table!.Columns!);
        Assert.DoesNotContain("Amount", updatedInfo.Table.Columns);
    }

    /// <summary>
    /// Tests renaming a numeric column to another numeric name.
    /// LLM use case: "rename column 60 to 120"
    /// Regression test for: Column names can be numeric (e.g. 60 for 60 months)
    /// </summary>
    [Fact]
    public async Task RenameColumn_NumericToNumeric_RenamesSuccessfully()
    {
        var testFile = await CreateTestFileWithTableAsync(nameof(RenameColumn_NumericToNumeric_RenamesSuccessfully));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // First add a numeric column
        await _tableCommands.AddColumnAsync(batch, "SalesTable", "60");

        // Then rename it to another numeric name
        var result = await _tableCommands.RenameColumnAsync(batch, "SalesTable", "60", "120");
        Assert.True(result.Success, $"Failed to rename numeric column to numeric name: {result.ErrorMessage}");

        // Verify column renamed
        var updatedInfo = await _tableCommands.GetAsync(batch, "SalesTable");
        Assert.Contains("120", updatedInfo.Table!.Columns!);
        Assert.DoesNotContain("60", updatedInfo.Table.Columns);
    }

    #endregion
}

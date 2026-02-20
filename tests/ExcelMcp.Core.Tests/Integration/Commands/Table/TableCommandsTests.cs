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
/// Modification tests use CreateModificationTestFile() for isolated files.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Medium")]
public partial class TableCommandsTests : IClassFixture<TableTestsFixture>
{
    private readonly TableCommands _tableCommands;
    private readonly IRangeCommands _rangeCommands;
    private readonly TableTestsFixture _fixture;
    private readonly string _tableFile;
    private readonly TableCreationResult _creationResult;

    /// <summary>
    /// Initializes a new instance of the <see cref="TableCommandsTests"/> class.
    /// </summary>
    public TableCommandsTests(TableTestsFixture fixture)
    {
        _tableCommands = new TableCommands();
        _rangeCommands = new RangeCommands();
        _fixture = fixture;
        _tableFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
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
    public void List_WithValidFile_ReturnsTables()
    {
        using var batch = ExcelSession.BeginBatch(_tableFile);
        var result = _tableCommands.List(batch);

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
    }

    /// <summary>
    /// Tests getting table details.
    /// LLM use case: "show me information about this table"
    /// </summary>
    [Fact]
    public void Info_WithValidTable_ReturnsTableDetails()
    {
        using var batch = ExcelSession.BeginBatch(_tableFile);
        var result = _tableCommands.Read(batch, "SalesTable");

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
    public void Create_WithValidData_CreatesTable()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Add data to a new location (different from the SalesTable created by fixture method)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Range["F1"].Value2 = "Name";
            sheet.Range["G1"].Value2 = "Value";
            sheet.Range["F2"].Value2 = "Test1";
            sheet.Range["G2"].Value2 = 100;
            return 0;
        });

        // Create table
        _tableCommands.Create(batch, "Sales", "TestTable", "F1:G2", true, "TableStyleLight1");
        // Create throws on error, so reaching here means success

        // Verify table was created
        var listResult = _tableCommands.List(batch);
        Assert.Contains(listResult.Tables, t => t.Name == "TestTable");
    }

    /// <summary>
    /// Tests deleting a table.
    /// LLM use case: "delete this table"
    /// </summary>
    [Fact]
    public void Delete_WithExistingTable_RemovesTable()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        _tableCommands.Delete(batch, "SalesTable");
        // Delete throws on error, so reaching here means success

        // Verify deletion
        var listResult = _tableCommands.List(batch);
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
    }

    /// <summary>
    /// Tests renaming a table.
    /// LLM use case: "rename this table"
    /// </summary>
    [Fact]
    public void Rename_WithExistingTable_RenamesSuccessfully()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        _tableCommands.Rename(batch, "SalesTable", "RevenueTable");
        // Rename throws on error, so reaching here means success

        // Verify rename
        var listResult = _tableCommands.List(batch);
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
        Assert.Contains(listResult.Tables, t => t.Name == "RevenueTable");
    }

    /// <summary>
    /// Tests resizing a table.
    /// LLM use case: "expand this table to include more rows"
    /// </summary>
    [Fact]
    public void Resize_WithExistingTable_ResizesSuccessfully()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        var initialInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.True(initialInfo.Success);

        _tableCommands.Resize(batch, "SalesTable", "A1:D10");
        // Resize throws on error, so reaching here means success

        // Verify resize
        var resizedInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.Equal(9, resizedInfo.Table!.RowCount); // 10 rows - 1 header
    }

    #endregion

    #region Column Operations (2 tests)

    /// <summary>
    /// Tests adding a column to a table.
    /// LLM use case: "add a new column to this table"
    /// </summary>
    [Fact]
    public void AddColumn_WithExistingTable_AddsColumnSuccessfully()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        var initialInfo = _tableCommands.Read(batch, "SalesTable");
        var initialColumnCount = initialInfo.Table!.Columns!.Count;

        _tableCommands.AddColumn(batch, "SalesTable", "NewColumn");
        // AddColumn throws on error, so reaching here means success

        // Verify column added
        var updatedInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.Equal(initialColumnCount + 1, updatedInfo.Table!.Columns!.Count);
        Assert.Contains("NewColumn", updatedInfo.Table.Columns);
    }

    /// <summary>
    /// Tests renaming a column in a table.
    /// LLM use case: "rename this table column"
    /// </summary>
    [Fact]
    public void RenameColumn_WithExistingColumn_RenamesSuccessfully()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        _tableCommands.RenameColumn(batch, "SalesTable", "Amount", "Revenue");
        // RenameColumn throws on error, so reaching here means success

        // Verify rename
        var info = _tableCommands.Read(batch, "SalesTable");
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
    public void Append_WithNewData_AddsRowsToTable()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        var newRows = new List<List<object?>>
        {
            new() { "West", "Widget", 500, DateTime.Now },
            new() { "East", "Gadget", 600, DateTime.Now }
        };

        _tableCommands.Append(batch, "SalesTable", newRows);
        // Append throws on error, so reaching here means success

        // Verify rows added
        var info = _tableCommands.Read(batch, "SalesTable");
        Assert.True(info.Table!.RowCount >= 6); // Original 4 + appended 2
    }

    /// <summary>
    /// Tests retrieving table data without filters.
    /// LLM use case: "read the table data for analysis"
    /// </summary>
    [Fact]
    public void GetData_WithoutFilters_ReturnsAllRows()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        var result = _tableCommands.GetData(batch, "SalesTable", visibleOnly: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("SalesTable", result.TableName);
        Assert.Equal(4, result.Headers.Count);
        Assert.Equal(4, result.RowCount); // Fixture data has 4 rows
        Assert.Equal(result.RowCount, result.Data.Count);
    }

    /// <summary>
    /// Tests retrieving only visible table rows after applying a filter.
    /// LLM use case: "get the filtered dataset"
    /// </summary>
    [Fact]
    public void GetData_WithVisibleOnlyFilter_ReturnsFilteredRows()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Apply filter so only North region remains visible
        _tableCommands.ApplyFilterValues(batch, "SalesTable", "Region", ["North"]);

        var result = _tableCommands.GetData(batch, "SalesTable", visibleOnly: true);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.RowCount);
        Assert.Single(result.Data);
        Assert.Equal("North", result.Data[0][0]?.ToString());
    }

    /// <summary>
    /// Tests getting structured reference for a table column.
    /// LLM use case: "get the structured reference formula for this table column"
    /// </summary>
    [Fact]
    public void GetStructuredReference_WithValidTable_ReturnsReference()
    {
        using var batch = ExcelSession.BeginBatch(_tableFile);
        var result = _tableCommands.GetStructuredReference(batch, "SalesTable", TableRegion.Data, "Amount");

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
    public void ApplyFilter_WithColumnCriteria_FiltersTable()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        _tableCommands.ApplyFilterValues(batch, "SalesTable", "Region", ["North"]);
        // ApplyFilter throws on error, so reaching here means success
    }

    /// <summary>
    /// Tests clearing all filters from a table.
    /// LLM use case: "remove all filters from this table"
    /// </summary>
    [Fact]
    public void ClearFilters_AfterFiltering_RemovesAllFilters()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Apply filter first
        _tableCommands.ApplyFilterValues(batch, "SalesTable", "Region", ["North"]);

        // Clear filters
        _tableCommands.ClearFilters(batch, "SalesTable");
        // ClearFilters throws on error, so reaching here means success
    }

    #endregion

    #region Totals Operations (2 tests)

    /// <summary>
    /// Tests enabling totals row on a table.
    /// LLM use case: "add a totals row to this table"
    /// </summary>
    [Fact]
    public void ToggleTotals_EnableTotals_AddsTotalsRow()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);
        _tableCommands.ToggleTotals(batch, "SalesTable", true);
        // ToggleTotals throws on error, so reaching here means success

        // Verify totals enabled
        var info = _tableCommands.Read(batch, "SalesTable");
        Assert.True(info.Table!.ShowTotals);
    }

    /// <summary>
    /// Tests setting a total function on a column.
    /// LLM use case: "set the total for this column to sum"
    /// </summary>
    [Fact]
    public void SetColumnTotal_WithSumFunction_SetsTotalFormula()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Enable totals first
        _tableCommands.ToggleTotals(batch, "SalesTable", true);

        // Set sum for Amount column
        _tableCommands.SetColumnTotal(batch, "SalesTable", "Amount", "Sum");
        // SetColumnTotal throws on error, so reaching here means success
    }

    #endregion

    #region Numeric Column Name Tests (3 tests)

    /// <summary>
    /// Tests adding a column with a purely numeric name.
    /// LLM use case: "add a column named 60 for 60 months data"
    /// Regression test for: Column names can be numeric (e.g. 60 for 60 months)
    /// </summary>
    [Fact]
    public void AddColumn_WithNumericName_AddsColumnSuccessfully()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        var initialInfo = _tableCommands.Read(batch, "SalesTable");
        var initialColumnCount = initialInfo.Table!.Columns!.Count;

        // Add column with purely numeric name
        _tableCommands.AddColumn(batch, "SalesTable", "60");
        // AddColumn throws on error, so reaching here means success

        // Verify column added
        var updatedInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.Equal(initialColumnCount + 1, updatedInfo.Table!.Columns!.Count);
        Assert.Contains("60", updatedInfo.Table.Columns);
    }

    /// <summary>
    /// Tests renaming a column to a purely numeric name.
    /// LLM use case: "rename this column to 12 for 12 months"
    /// Regression test for: Column names can be numeric (e.g. 60 for 60 months)
    /// </summary>
    [Fact]
    public void RenameColumn_ToNumericName_RenamesSuccessfully()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Rename "Amount" column to numeric name "60"
        _tableCommands.RenameColumn(batch, "SalesTable", "Amount", "60");
        // RenameColumn throws on error, so reaching here means success

        // Verify column renamed
        var updatedInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.Contains("60", updatedInfo.Table!.Columns!);
        Assert.DoesNotContain("Amount", updatedInfo.Table.Columns);
    }

    /// <summary>
    /// Tests renaming a numeric column to another numeric name.
    /// LLM use case: "rename column 60 to 120"
    /// Regression test for: Column names can be numeric (e.g. 60 for 60 months)
    /// </summary>
    [Fact]
    public void RenameColumn_NumericToNumeric_RenamesSuccessfully()
    {
        var testFile = _fixture.CreateModificationTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // First add a numeric column
        _tableCommands.AddColumn(batch, "SalesTable", "60");

        // Then rename it to another numeric name
        _tableCommands.RenameColumn(batch, "SalesTable", "60", "120");
        // RenameColumn throws on error, so reaching here means success

        // Verify column renamed
        var updatedInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.Contains("120", updatedInfo.Table!.Columns!);
        Assert.DoesNotContain("60", updatedInfo.Table.Columns);
    }

    #endregion
}





using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Fact]
    public void List_WithValidFile_ReturnsTables()
    {
        var batch = _fixture.BatchToken;
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
        var batch = _fixture.BatchToken;
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

        var batch = _fixture.BatchToken;

        // Add data to a new location (different from SalesTable).
        _rangeCommands.SetValues(
            batch,
            "Sales",
            "F1:G2",
            [["Name", "Value"], ["Test1", 100]]);

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

        var batch = _fixture.BatchToken;
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

        var batch = _fixture.BatchToken;
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

        var batch = _fixture.BatchToken;

        var initialInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.True(initialInfo.Success);

        _tableCommands.Resize(batch, "SalesTable", "A1:D10");
        // Resize throws on error, so reaching here means success

        // Verify resize
        var resizedInfo = _tableCommands.Read(batch, "SalesTable");
        Assert.Equal(9, resizedInfo.Table!.RowCount); // 10 rows - 1 header
    }

    /// <summary>
    /// Tests adding a column to a table.
    /// LLM use case: "add a new column to this table"
    /// </summary>
    [Fact]
    public void AddColumn_WithExistingTable_AddsColumnSuccessfully()
    {

        var batch = _fixture.BatchToken;

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

        var batch = _fixture.BatchToken;

        _tableCommands.RenameColumn(batch, "SalesTable", "Amount", "Revenue");
        // RenameColumn throws on error, so reaching here means success

        // Verify rename
        var info = _tableCommands.Read(batch, "SalesTable");
        Assert.Contains("Revenue", info.Table!.Columns!);
        Assert.DoesNotContain("Amount", info.Table.Columns);
    }

    /// <summary>
    /// Tests appending rows to a table.
    /// LLM use case: "add these rows to the table"
    /// </summary>
    [Fact]
    public void Append_WithNewData_AddsRowsToTable()
    {

        var batch = _fixture.BatchToken;

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

        var batch = _fixture.BatchToken;

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

        var batch = _fixture.BatchToken;

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
        var batch = _fixture.BatchToken;
        var result = _tableCommands.GetStructuredReference(batch, "SalesTable", TableRegion.Data, "Amount");

        Assert.True(result.Success);
        Assert.NotNull(result.StructuredReference);
        Assert.Contains("SalesTable", result.StructuredReference);
        Assert.Contains("Amount", result.StructuredReference);
    }

    /// <summary>
    /// Tests applying a filter to a table column.
    /// LLM use case: "filter this table to show only these values"
    /// </summary>
    [Fact]
    public void ApplyFilter_WithColumnCriteria_FiltersTable()
    {

        var batch = _fixture.BatchToken;
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

        var batch = _fixture.BatchToken;

        // Apply filter first
        _tableCommands.ApplyFilterValues(batch, "SalesTable", "Region", ["North"]);

        // Clear filters
        _tableCommands.ClearFilters(batch, "SalesTable");
        // ClearFilters throws on error, so reaching here means success
    }

    /// <summary>
    /// Tests enabling totals row on a table.
    /// LLM use case: "add a totals row to this table"
    /// </summary>
    [Fact]
    public void ToggleTotals_EnableTotals_AddsTotalsRow()
    {

        var batch = _fixture.BatchToken;
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

        var batch = _fixture.BatchToken;

        // Enable totals first
        _tableCommands.ToggleTotals(batch, "SalesTable", true);

        // Set sum for Amount column
        _tableCommands.SetColumnTotal(batch, "SalesTable", "Amount", "Sum");
        // SetColumnTotal throws on error, so reaching here means success
    }

    /// <summary>
    /// Tests adding a column with a purely numeric name.
    /// LLM use case: "add a column named 60 for 60 months data"
    /// Regression test for: Column names can be numeric (e.g. 60 for 60 months)
    /// </summary>
    [Fact]
    public void AddColumn_WithNumericName_AddsColumnSuccessfully()
    {

        var batch = _fixture.BatchToken;

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

        var batch = _fixture.BatchToken;

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

        var batch = _fixture.BatchToken;

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
}

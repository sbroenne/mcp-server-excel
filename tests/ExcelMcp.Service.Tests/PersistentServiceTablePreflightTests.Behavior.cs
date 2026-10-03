using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Filtering;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Fact]
    public void List_WithValidFile_ReturnsTables()
    {
        var batch = _fixture.BatchToken;
        var result = _tableCommands.List(batch);

        RequireSuccess(result);
        var table = Assert.Single(result.Tables);
        Assert.Equal("SalesTable", table.Name);
        AssertSalesTableInfo(table);
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

        RequireSuccess(result);
        Assert.NotNull(result.Table);
        Assert.Equal("SalesTable", result.Table.Name);
        Assert.Equal("Sales", result.Table.SheetName);
        Assert.True(result.Table.HasHeaders);
        Assert.Equal(4, result.Table.Columns?.Count);
        AssertSalesTableInfo(result.Table);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
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
        RequireSuccess(_rangeCommands.SetValues(
            batch,
            "Sales",
            "F1:G2",
            [["Name", "Value"], ["Test1", 100]]));

        // Create table
        RequireSuccess(_tableCommands.Create(batch, "Sales", "TestTable", "F1:G2", true, "TableStyleLight1"));

        // Verify table was created
        var listResult = RequireSuccess(_tableCommands.List(batch));
        Assert.Contains(listResult.Tables, t => t.Name == "TestTable");
        Assert.Equal(2, listResult.Tables.Count);
        var info = RequireSuccess(_tableCommands.Read(batch, "TestTable")).Table;
        Assert.NotNull(info);
        Assert.Equal("$F$1:$G$2", info.Range);
        Assert.Equal("TableStyleLight1", info.TableStyle);
        Assert.Equal(["Name", "Value"], info.Columns);
        var data = RequireSuccess(_tableCommands.GetData(batch, "TestTable"));
        AssertNamedAmountRow(Assert.Single(data.Data), "Test1", 100);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
    }

    /// <summary>
    /// Tests deleting a table.
    /// LLM use case: "delete this table"
    /// </summary>
    [Fact]
    public void Delete_WithExistingTable_RemovesTable()
    {

        var batch = _fixture.BatchToken;
        RequireSuccess(_tableCommands.Delete(batch, "SalesTable"));

        // Verify deletion
        var listResult = RequireSuccess(_tableCommands.List(batch));
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
        Assert.Empty(listResult.Tables);
        var cells = RequireSuccess(_rangeCommands.GetValues(batch, "Sales", "A1:D5"));
        Assert.Equal(["Region", "Product", "Amount", "Date"], cells.Values[0]);
        Assert.Equal(["North", "South", "East", "West"], cells.Values.Skip(1).Select(row => row[0]));
        Assert.Equal(["Widget", "Gadget", "Widget", "Gadget"], cells.Values.Skip(1).Select(row => row[1]));
        Assert.Equal([100d, 250d, 150d, 300d], cells.Values.Skip(1).Select(row => Convert.ToDouble(row[2], CultureInfo.InvariantCulture)));
        Assert.Equal([new DateTime(2025, 1, 15), new DateTime(2025, 2, 20),
            new DateTime(2025, 3, 10), new DateTime(2025, 1, 25)],
            cells.Values.Skip(1).Select(row => DateTime.FromOADate(Convert.ToDouble(row[3], CultureInfo.InvariantCulture))));
    }

    /// <summary>
    /// Tests renaming a table.
    /// LLM use case: "rename this table"
    /// </summary>
    [Fact]
    public void Rename_WithExistingTable_RenamesSuccessfully()
    {

        var batch = _fixture.BatchToken;
        RequireSuccess(_tableCommands.Rename(batch, "SalesTable", "RevenueTable"));

        // Verify rename
        var listResult = RequireSuccess(_tableCommands.List(batch));
        Assert.DoesNotContain(listResult.Tables, t => t.Name == "SalesTable");
        Assert.Contains(listResult.Tables, t => t.Name == "RevenueTable");
        Assert.Equal("RevenueTable", Assert.Single(listResult.Tables).Name);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "RevenueTable")));
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
        RequireSuccess(initialInfo);

        RequireSuccess(_tableCommands.Resize(batch, "SalesTable", "A1:D10"));

        // Verify resize
        var resizedInfo = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.Equal(9, resizedInfo.Table!.RowCount); // 10 rows - 1 header
        Assert.Equal("$A$1:$D$10", resizedInfo.Table.Range);
        var data = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));
        AssertSalesRows(data.Data.Take(4).ToList());
        Assert.Equal(9, data.Data.Count);
        Assert.All(data.Data.Skip(4), row => Assert.All(row, Assert.Null));
    }

    /// <summary>
    /// Tests adding a column to a table.
    /// LLM use case: "add a new column to this table"
    /// </summary>
    [Fact]
    public void AddColumn_WithExistingTable_AddsColumnSuccessfully()
    {

        var batch = _fixture.BatchToken;

        var initialInfo = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        var initialColumnCount = initialInfo.Table!.Columns!.Count;

        RequireSuccess(_tableCommands.AddColumn(batch, "SalesTable", "NewColumn"));

        // Verify column added
        var updatedInfo = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.Equal(initialColumnCount + 1, updatedInfo.Table!.Columns!.Count);
        Assert.Contains("NewColumn", updatedInfo.Table.Columns);
        Assert.Equal(["Region", "Product", "Amount", "Date", "NewColumn"], updatedInfo.Table.Columns);
        var data = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));
        AssertSalesRows(data.Data.Select(row => row.Take(4).ToList()).ToList());
        Assert.All(data.Data, row => Assert.Null(row[4]));
    }

    /// <summary>
    /// Tests renaming a column in a table.
    /// LLM use case: "rename this table column"
    /// </summary>
    [Fact]
    public void RenameColumn_WithExistingColumn_RenamesSuccessfully()
    {

        var batch = _fixture.BatchToken;

        RequireSuccess(_tableCommands.RenameColumn(batch, "SalesTable", "Amount", "Revenue"));

        // Verify rename
        var info = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.Contains("Revenue", info.Table!.Columns!);
        Assert.DoesNotContain("Amount", info.Table.Columns);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")), ["Region", "Product", "Revenue", "Date"]);
    }

    /// <summary>
    /// Tests appending rows to a table.
    /// LLM use case: "add these rows to the table"
    /// </summary>
    [Fact]
    public void Append_WithNewData_AddsRowsToTable()
    {

        var batch = _fixture.BatchToken;

        var original = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));
        AssertSalesData(original);
        var appendedDate = new DateTime(2025, 4, 1);
        var newRows = new List<List<object?>>
        {
            new() { "West", "Widget", 500, appendedDate },
            new() { "East", "Gadget", 600, appendedDate }
        };

        RequireSuccess(_tableCommands.Append(batch, "SalesTable", newRows));

        var info = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.Equal(original.RowCount + 2, info.Table!.RowCount);
        var actual = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));
        Assert.Equal(original.Headers, actual.Headers);
        Assert.Equal(original.RowCount + 2, actual.RowCount);
        Assert.Equal(JsonSerializer.Serialize(original.Data),
            JsonSerializer.Serialize(actual.Data.Take(original.RowCount)));
        for (var index = 0; index < newRows.Count; index++)
        {
            var row = actual.Data[original.RowCount + index];
            Assert.Equal(4, row.Count);
            Assert.Equal(newRows[index][0], row[0]);
            Assert.Equal(newRows[index][1], row[1]);
            Assert.Equal(Convert.ToDouble(newRows[index][2], CultureInfo.InvariantCulture),
                Convert.ToDouble(row[2], CultureInfo.InvariantCulture));
            Assert.Equal(appendedDate, DateTime.Parse(
                Assert.IsType<string>(row[3]), CultureInfo.InvariantCulture));
        }
    }

    [Theory]
    [InlineData(3, false)]
    [InlineData(5, false)]
    [InlineData(3, true)]
    [InlineData(5, true)]
    public void Append_WithMismatchedLaterRow_RejectsBeforeWriting(int secondRowColumns, bool fromFile)
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F6:F7", [["Retained"], [42]]);
        var before = GetSourceState("Sales", "A1:F8");
        var calculation = _fixture.ExecuteRawVerification((context, _) => context.App.Calculation);
        List<List<object?>> rows =
        [
            ["North", "Widget", 500, "2025-04-01"],
            Enumerable.Range(0, secondRowColumns).Select(index => (object?)index).ToList()
        ];

        var file = fromFile ? _fixture.CreateInputFile(".json", JsonSerializer.Serialize(rows)) : null;
        var error = Assert.Throws<ArgumentException>(() =>
            _tableCommands.Append(batch, "SalesTable", fromFile ? null : rows, file));

        Assert.Contains("row 2", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("4 columns", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertPreflightPreserved(before, "A1:F8");
        Assert.Equal(calculation, _fixture.ExecuteRawVerification((context, _) => context.App.Calculation));
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

        RequireSuccess(result);
        Assert.Equal("SalesTable", result.TableName);
        Assert.Equal(4, result.Headers.Count);
        Assert.Equal(4, result.RowCount); // Fixture data has 4 rows
        Assert.Equal(result.RowCount, result.Data.Count);
        AssertSalesData(result);
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
        RequireSuccess(_tableCommands.ApplyFilter(batch, "SalesTable", "Region",
            new FilterOptions { FilterOperator = FilterOperator.Values, Values = ["North"] }));

        var result = _tableCommands.GetData(batch, "SalesTable", visibleOnly: true);

        RequireSuccess(result);
        Assert.Equal(1, result.RowCount);
        Assert.Single(result.Data);
        Assert.Equal("North", result.Data[0][0]?.ToString());
        AssertSalesRows(result.Data, [0]);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
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

        RequireSuccess(result);
        Assert.NotNull(result.StructuredReference);
        Assert.Equal("SalesTable[[Amount]]", result.StructuredReference);
    }

    /// <summary>
    /// Tests applying a filter to a table column.
    /// LLM use case: "filter this table to show only these values"
    /// </summary>
    [Fact]
    public void ApplyFilter_WithColumnCriteria_FiltersTable()
    {

        var batch = _fixture.BatchToken;
        var before = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));
        RequireSuccess(_tableCommands.ApplyFilter(batch, "SalesTable", "Region",
            new FilterOptions { FilterOperator = FilterOperator.Values, Values = ["North"] }));
        var visible = RequireSuccess(_tableCommands.GetData(batch, "SalesTable", visibleOnly: true));
        AssertSalesRows(visible.Data, [0]);
        Assert.Equal(JsonSerializer.Serialize(before.Data),
            JsonSerializer.Serialize(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")).Data));
    }

    /// <summary>
    /// Tests clearing all filters from a table.
    /// LLM use case: "remove all filters from this table"
    /// </summary>
    [Fact]
    public void ClearFilters_AfterFiltering_RemovesAllFilters()
    {

        var batch = _fixture.BatchToken;
        var before = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));

        // Apply filter first
        RequireSuccess(_tableCommands.ApplyFilter(batch, "SalesTable", "Region",
            new FilterOptions { FilterOperator = FilterOperator.Values, Values = ["North"] }));
        Assert.Single(RequireSuccess(_tableCommands.GetData(batch, "SalesTable", visibleOnly: true)).Data);

        // Clear filters
        RequireSuccess(_tableCommands.ClearFilters(batch, "SalesTable"));
        var restored = RequireSuccess(_tableCommands.GetData(batch, "SalesTable", visibleOnly: true));
        Assert.Equal(JsonSerializer.Serialize(before.Data), JsonSerializer.Serialize(restored.Data));
        Assert.False(RequireSuccess(_tableCommands.GetFilters(batch, "SalesTable")).HasActiveFilters);
        AssertSalesData(restored);
    }

    /// <summary>
    /// Tests enabling totals row on a table.
    /// LLM use case: "add a totals row to this table"
    /// </summary>
    [Fact]
    public void ToggleTotals_EnableTotals_AddsTotalsRow()
    {

        var batch = _fixture.BatchToken;
        RequireSuccess(_tableCommands.ToggleTotals(batch, "SalesTable", true));

        // Verify totals enabled
        var info = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.True(info.Table!.ShowTotals);
        Assert.Equal("$A$1:$D$6", info.Table.Range);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
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
        RequireSuccess(_tableCommands.ToggleTotals(batch, "SalesTable", true));

        // Set sum for Amount column
        RequireSuccess(_tableCommands.SetColumnTotal(batch, "SalesTable", "Amount", "Sum"));
        var formula = RequireSuccess(_rangeCommands.GetFormulas(batch, "Sales", "C6"));
        Assert.Equal("=SUBTOTAL(109,[Amount])", Assert.Single(Assert.Single(formula.Formulas)));
        var values = RequireSuccess(_rangeCommands.GetValues(batch, "Sales", "C6"));
        Assert.Equal(800, Convert.ToDouble(Assert.Single(Assert.Single(values.Values)), CultureInfo.InvariantCulture));
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
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

        var initialInfo = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        var initialColumnCount = initialInfo.Table!.Columns!.Count;

        // Add column with purely numeric name
        RequireSuccess(_tableCommands.AddColumn(batch, "SalesTable", "60"));

        // Verify column added
        var updatedInfo = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.Equal(initialColumnCount + 1, updatedInfo.Table!.Columns!.Count);
        Assert.Contains("60", updatedInfo.Table.Columns);
        Assert.Equal(["Region", "Product", "Amount", "Date", "60"], updatedInfo.Table.Columns);
        var data = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));
        AssertSalesRows(data.Data.Select(row => row.Take(4).ToList()).ToList());
        Assert.All(data.Data, row => Assert.Null(row[4]));
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
        RequireSuccess(_tableCommands.RenameColumn(batch, "SalesTable", "Amount", "60"));

        // Verify column renamed
        var updatedInfo = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.Contains("60", updatedInfo.Table!.Columns!);
        Assert.DoesNotContain("Amount", updatedInfo.Table.Columns);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")), ["Region", "Product", "60", "Date"]);
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
        RequireSuccess(_tableCommands.AddColumn(batch, "SalesTable", "60"));

        // Then rename it to another numeric name
        RequireSuccess(_tableCommands.RenameColumn(batch, "SalesTable", "60", "120"));

        // Verify column renamed
        var updatedInfo = RequireSuccess(_tableCommands.Read(batch, "SalesTable"));
        Assert.Contains("120", updatedInfo.Table!.Columns!);
        Assert.DoesNotContain("60", updatedInfo.Table.Columns);
        Assert.Equal(["Region", "Product", "Amount", "Date", "120"], updatedInfo.Table.Columns);
        var data = RequireSuccess(_tableCommands.GetData(batch, "SalesTable"));
        AssertSalesRows(data.Data.Select(row => row.Take(4).ToList()).ToList());
        Assert.All(data.Data, row => Assert.Null(row[4]));
    }

    private static void AssertSalesTableInfo(TableInfo table)
    {
        Assert.Equal("Sales", table.SheetName);
        Assert.Equal("$A$1:$D$5", table.Range);
        Assert.True(table.HasHeaders);
        Assert.False(table.ShowTotals);
        Assert.Equal(4, table.RowCount);
        Assert.Equal(4, table.ColumnCount);
        Assert.Equal(["Region", "Product", "Amount", "Date"], table.Columns);
        Assert.Equal("TableStyleMedium2", table.TableStyle);
    }

    private static void AssertSalesData(TableDataResult result, string[]? expectedHeaders = null)
    {
        RequireSuccess(result);
        Assert.Equal(expectedHeaders ?? ["Region", "Product", "Amount", "Date"], result.Headers);
        Assert.Equal(4, result.RowCount);
        Assert.Equal(4, result.ColumnCount);
        AssertSalesRows(result.Data);
    }

    private static void AssertSalesRows(List<List<object?>> actual, int[]? indexes = null)
    {
        List<List<object?>> expected =
        [
            ["North", "Widget", 100, new DateTime(2025, 1, 15)],
            ["South", "Gadget", 250, new DateTime(2025, 2, 20)],
            ["East", "Widget", 150, new DateTime(2025, 3, 10)],
            ["West", "Gadget", 300, new DateTime(2025, 1, 25)]
        ];
        var selected = (indexes ?? [0, 1, 2, 3]).Select(index => expected[index]).ToList();
        Assert.Equal(selected.Count, actual.Count);
        for (var index = 0; index < selected.Count; index++)
        {
            Assert.Equal(4, actual[index].Count);
            Assert.Equal(selected[index][0], actual[index][0]);
            Assert.Equal(selected[index][1], actual[index][1]);
            Assert.Equal(Convert.ToDouble(selected[index][2], CultureInfo.InvariantCulture),
                Convert.ToDouble(actual[index][2], CultureInfo.InvariantCulture));
            Assert.Equal(Assert.IsType<DateTime>(selected[index][3]),
                DateTime.FromOADate(Convert.ToDouble(actual[index][3], CultureInfo.InvariantCulture)));
        }
    }

    private static void AssertNamedAmountRow(List<object?> actual, string label, double amount)
    {
        Assert.Equal(2, actual.Count);
        Assert.Equal(label, actual[0]);
        Assert.Equal(amount, Convert.ToDouble(actual[1], CultureInfo.InvariantCulture));
    }
}

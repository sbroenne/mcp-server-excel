using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for DAX EVALUATE query execution.
/// Tests verify that DAX EVALUATE queries can be executed against the Data Model
/// and return tabular results via the ADO connection.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class PersistentServiceDataModelEvaluateTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IDataModelCommands _dataModelCommands =
        fixture.CreateCommands<IDataModelCommands>();
    private static readonly (int Id, double Date, int Customer, int Product, decimal Amount, int Quantity)[] ExpectedSales =
    [
        (1, 45306, 101, 1001, 150, 2),
        (2, 45311, 102, 1002, 250, 3),
        (3, 45332, 101, 1003, 175, 1),
        (4, 45337, 103, 1001, 300, 4),
        (5, 45356, 102, 1002, 125, 2),
        (6, 45361, 104, 1003, 450, 5),
        (7, 45394, 101, 1001, 200, 2),
        (8, 45400, 103, 1002, 350, 4),
        (9, 45420, 105, 1003, 275, 3),
        (10, 45434, 102, 1001, 180, 2)
    ];

    #region Basic EVALUATE Tests

    /// <summary>
    /// Tests that a simple EVALUATE query returns table data.
    /// LLM use case: "show me all rows from this Data Model table"
    /// </summary>
    [Fact]
    public void Evaluate_SimpleTableQuery_ReturnsRows()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE 'SalesTable' ORDER BY 'SalesTable'[SalesID]");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        Assert.True(result.RowCount > 0, "Expected at least one row");
        Assert.True(result.ColumnCount > 0, "Expected at least one column");

        // Column names include table prefix (e.g., "SalesTable[CustomerID]")
        Assert.True(result.Columns.Any(c => c.Contains("CustomerID", StringComparison.OrdinalIgnoreCase)),
            $"Expected a column containing 'CustomerID', got: {string.Join(", ", result.Columns)}");
        Assert.True(result.Columns.Any(c => c.Contains("Amount", StringComparison.OrdinalIgnoreCase)),
            $"Expected a column containing 'Amount', got: {string.Join(", ", result.Columns)}");
        AssertSalesRows(result, ExpectedSales);
    }

    /// <summary>
    /// Tests EVALUATE with SUMMARIZE for aggregated results.
    /// LLM use case: "summarize sales by customer"
    /// </summary>
    [Fact]
    public void Evaluate_SummarizeQuery_ReturnsAggregatedData()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE SUMMARIZE('SalesTable', 'SalesTable'[CustomerID], \"TotalAmount\", SUM('SalesTable'[Amount])) ORDER BY 'SalesTable'[CustomerID]");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        Assert.True(result.RowCount > 0, "Expected aggregated rows");

        // Column names include table prefix
        Assert.True(result.Columns.Any(c => c.Contains("CustomerID", StringComparison.OrdinalIgnoreCase)),
            $"Expected a column containing 'CustomerID', got: {string.Join(", ", result.Columns)}");
        Assert.True(result.Columns.Any(c => c.Contains("Amount", StringComparison.OrdinalIgnoreCase) ||
                                           c.Contains("TotalAmount", StringComparison.OrdinalIgnoreCase)),
            $"Expected a column containing 'Amount' or 'TotalAmount', got: {string.Join(", ", result.Columns)}");
        Assert.Equal(2, result.ColumnCount);
        var expected = ExpectedSales.GroupBy(row => row.Customer).OrderBy(group => group.Key).ToArray();
        Assert.Equal(expected.Length, result.RowCount);
        Assert.Equal(expected.Length, result.Rows.Count);
        for (var index = 0; index < expected.Length; index++)
        {
            Assert.Equal(expected[index].Key, Convert.ToInt32(result.Rows[index][0],
                System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(expected[index].Sum(row => row.Amount), Convert.ToDecimal(result.Rows[index][1],
                System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    /// <summary>
    /// Tests EVALUATE with FILTER for filtered results.
    /// LLM use case: "show me sales greater than 100"
    /// </summary>
    [Fact]
    public void Evaluate_FilterQuery_ReturnsFilteredRows()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE FILTER('SalesTable', 'SalesTable'[Amount] > 200) ORDER BY 'SalesTable'[SalesID]");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);

        AssertSalesRows(result, ExpectedSales.Where(row => row.Amount > 200).ToArray());
    }

    /// <summary>
    /// Tests EVALUATE with ROW for scalar results.
    /// LLM use case: "calculate total sales"
    /// </summary>
    [Fact]
    public void Evaluate_RowQuery_ReturnsSingleRow()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE ROW(\"Probe\", 42)");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);
        Assert.Equal(1, result.RowCount); // ROW returns exactly one row

        // Should have one column with the computed value
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal("[Probe]", Assert.Single(result.Columns));
        Assert.Equal(42m, Convert.ToDecimal(Assert.Single(Assert.Single(result.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
    }

    #endregion

    #region Error Handling Tests

    /// <summary>
    /// Tests that null/empty query throws ArgumentException.
    /// </summary>
    [Fact]
    public void Evaluate_NullQuery_ThrowsArgumentException()
    {
        var batch = _fixture.BatchToken;

        var ex = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.Evaluate(batch, ""));

        Assert.Contains("daxQuery", ex.Message);
        AssertSalesRows(
            RequireSuccess(_dataModelCommands.Evaluate(
                batch, "EVALUATE 'SalesTable' ORDER BY 'SalesTable'[SalesID]")),
            ExpectedSales);
    }

    /// <summary>
    /// Tests that non-EVALUATE query returns error.
    /// (Only EVALUATE queries return tabular results)
    /// </summary>
    [Fact]
    public void Evaluate_NonEvaluateQuery_HandlesGracefully()
    {
        var batch = _fixture.BatchToken;

        var ex = Assert.Throws<InvalidOperationException>(() =>
            _dataModelCommands.Evaluate(batch, "DEFINE VAR x = 1"));

        Assert.Contains("ComInterop/", ex.Message, StringComparison.Ordinal);
        AssertSalesRows(_dataModelCommands.Evaluate(batch,
            "EVALUATE 'SalesTable' ORDER BY 'SalesTable'[SalesID]"), ExpectedSales);
    }

    #endregion

    #region Advanced Query Tests

    /// <summary>
    /// Tests EVALUATE with CALCULATETABLE for context-modified results.
    /// LLM use case: "show sales filtered by specific conditions"
    /// </summary>
    [Fact]
    public void Evaluate_CalculateTableQuery_ReturnsModifiedContext()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE CALCULATETABLE('SalesTable', 'SalesTable'[CustomerID] = 101) ORDER BY 'SalesTable'[SalesID]");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);

        AssertSalesRows(result, ExpectedSales.Where(row => row.Customer == 101).ToArray());
    }

    /// <summary>
    /// Tests EVALUATE with TOPN for limited results.
    /// LLM use case: "show me top 5 sales by amount"
    /// </summary>
    [Fact]
    public void Evaluate_TopNQuery_ReturnsLimitedRows()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE TOPN(3, 'SalesTable', 'SalesTable'[Amount], DESC) ORDER BY 'SalesTable'[Amount] DESC");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);
        AssertSalesRows(result, ExpectedSales.OrderByDescending(row => row.Amount).Take(3).ToArray());
    }

    /// <summary>
    /// Tests EVALUATE with DISTINCT for unique values.
    /// LLM use case: "show me unique customer IDs"
    /// </summary>
    [Fact]
    public void Evaluate_DistinctQuery_ReturnsUniqueValues()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE DISTINCT('SalesTable'[CustomerID]) ORDER BY 'SalesTable'[CustomerID]");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);
        Assert.Equal(1, result.ColumnCount); // DISTINCT on single column returns single column

        // Verify all values are unique
        var values = result.Rows.Select(r => r[0]).ToList();
        var uniqueValues = values.Distinct().ToList();
        Assert.Equal(values.Count, uniqueValues.Count);
        Assert.Equal(ExpectedSales.Select(row => row.Customer).Distinct().Order(),
            result.Rows.Select(row => Convert.ToInt32(Assert.Single(row),
                System.Globalization.CultureInfo.InvariantCulture)));
        Assert.Equal(5, result.RowCount);
    }

    #endregion

    private static void AssertSalesRows(
        DaxEvaluateResult result,
        (int Id, double Date, int Customer, int Product, decimal Amount, int Quantity)[] expected)
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage), result.ErrorMessage);
        Assert.Equal(expected.Length, result.RowCount);
        Assert.Equal(expected.Length, result.Rows.Count);
        Assert.Equal(6, result.ColumnCount);
        Assert.Equal(result.ColumnCount, result.Columns.Count);
        Assert.Equal(
            ["SalesTable[Amount]", "SalesTable[CustomerID]", "SalesTable[Date]",
                "SalesTable[ProductID]", "SalesTable[Quantity]", "SalesTable[SalesID]"],
            result.Columns.Order(StringComparer.Ordinal));
        var id = RequiredColumn("SalesID");
        var date = RequiredColumn("Date");
        var customer = RequiredColumn("CustomerID");
        var product = RequiredColumn("ProductID");
        var amount = RequiredColumn("Amount");
        var quantity = RequiredColumn("Quantity");
        for (var index = 0; index < expected.Length; index++)
        {
            var row = result.Rows[index];
            Assert.Equal(6, row.Count);
            Assert.NotNull(row[id]);
            Assert.NotNull(row[date]);
            Assert.NotNull(row[customer]);
            Assert.NotNull(row[product]);
            Assert.NotNull(row[amount]);
            Assert.NotNull(row[quantity]);
            Assert.Equal(expected[index].Id, Convert.ToInt32(row[id], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(expected[index].Date, Convert.ToDouble(
                row[date], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(expected[index].Customer, Convert.ToInt32(row[customer], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(expected[index].Product, Convert.ToInt32(row[product], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(expected[index].Amount, Convert.ToDecimal(row[amount], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(expected[index].Quantity, Convert.ToInt32(row[quantity], System.Globalization.CultureInfo.InvariantCulture));
        }

        int RequiredColumn(string name)
        {
            var index = result.Columns.FindIndex(column =>
                column.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                column.EndsWith($"[{name}]", StringComparison.OrdinalIgnoreCase));
            Assert.InRange(index, 0, result.ColumnCount - 1);
            return index;
        }
    }
}

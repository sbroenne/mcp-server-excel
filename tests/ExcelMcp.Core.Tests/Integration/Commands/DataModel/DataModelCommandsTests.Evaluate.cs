using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for DAX EVALUATE query execution.
/// Tests verify that DAX EVALUATE queries can be executed against the Data Model
/// and return tabular results via the ADO connection.
/// </summary>
[Collection("DataModel")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class DataModelCommandsTests_Evaluate
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly string _dataModelFile;

    public DataModelCommandsTests_Evaluate(DataModelPivotTableFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _dataModelFile = fixture.TestFilePath;
    }

    #region Basic EVALUATE Tests

    /// <summary>
    /// Tests that a simple EVALUATE query returns table data.
    /// LLM use case: "show me all rows from this Data Model table"
    /// </summary>
    [Fact]
    public void Evaluate_SimpleTableQuery_ReturnsRows()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Evaluate(batch, "EVALUATE 'SalesTable'");

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
    }

    /// <summary>
    /// Tests EVALUATE with SUMMARIZE for aggregated results.
    /// LLM use case: "summarize sales by customer"
    /// </summary>
    [Fact]
    public void Evaluate_SummarizeQuery_ReturnsAggregatedData()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE SUMMARIZE('SalesTable', 'SalesTable'[CustomerID], \"TotalAmount\", SUM('SalesTable'[Amount]))");

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
    }

    /// <summary>
    /// Tests EVALUATE with FILTER for filtered results.
    /// LLM use case: "show me sales greater than 100"
    /// </summary>
    [Fact]
    public void Evaluate_FilterQuery_ReturnsFilteredRows()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE FILTER('SalesTable', 'SalesTable'[Amount] > 100)");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);

        // All returned rows should have Amount > 100
        var amountColumnIndex = result.Columns.FindIndex(c =>
            c.Equals("Amount", StringComparison.OrdinalIgnoreCase) ||
            c.EndsWith("[Amount]", StringComparison.OrdinalIgnoreCase));

        if (amountColumnIndex >= 0 && result.Rows.Count > 0)
        {
            foreach (var row in result.Rows)
            {
                if (row[amountColumnIndex] != null)
                {
                    var amount = Convert.ToDecimal(row[amountColumnIndex], System.Globalization.CultureInfo.InvariantCulture);
                    Assert.True(amount > 100, $"Expected Amount > 100, got {amount}");
                }
            }
        }
    }

    /// <summary>
    /// Tests EVALUATE with ROW for scalar results.
    /// LLM use case: "calculate total sales"
    /// </summary>
    [Fact]
    public void Evaluate_RowQuery_ReturnsSingleRow()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE ROW(\"TotalSales\", SUM('SalesTable'[Amount]))");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);
        Assert.Equal(1, result.RowCount); // ROW returns exactly one row

        // Should have one column with the computed value
        Assert.Equal(1, result.ColumnCount);
    }

    #endregion

    #region Error Handling Tests

    /// <summary>
    /// Tests that invalid DAX query throws exception.
    /// LLM use case: handling syntax errors
    /// </summary>
    [Fact]
    public void Evaluate_InvalidDax_ThrowsException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Invalid DAX syntax - throws COMException from Excel
        var ex = Assert.ThrowsAny<Exception>(() =>
            _dataModelCommands.Evaluate(batch, "EVALUATE INVALID_FUNCTION()"));

        Assert.NotNull(ex);
    }

    /// <summary>
    /// Tests that null/empty query throws ArgumentException.
    /// </summary>
    [Fact]
    public void Evaluate_NullQuery_ThrowsArgumentException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.Evaluate(batch, ""));

        Assert.Contains("daxQuery", ex.Message);
    }

    /// <summary>
    /// Tests that non-EVALUATE query returns error.
    /// (Only EVALUATE queries return tabular results)
    /// </summary>
    [Fact]
    public void Evaluate_NonEvaluateQuery_HandlesGracefully()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // DEFINE without EVALUATE doesn't return data
        // This might throw or return empty depending on implementation
        var ex = Assert.ThrowsAny<Exception>(() =>
            _dataModelCommands.Evaluate(batch, "DEFINE VAR x = 1"));

        // Should indicate an error occurred
        Assert.NotNull(ex);
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
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE CALCULATETABLE('SalesTable', 'SalesTable'[CustomerID] = 1)");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);

        // All rows should have CustomerID = 1
        var customerIdIndex = result.Columns.FindIndex(c =>
            c.Equals("CustomerID", StringComparison.OrdinalIgnoreCase) ||
            c.EndsWith("[CustomerID]", StringComparison.OrdinalIgnoreCase));

        if (customerIdIndex >= 0 && result.Rows.Count > 0)
        {
            foreach (var row in result.Rows)
            {
                if (row[customerIdIndex] != null)
                {
                    var customerId = Convert.ToInt32(row[customerIdIndex], System.Globalization.CultureInfo.InvariantCulture);
                    Assert.Equal(1, customerId);
                }
            }
        }
    }

    /// <summary>
    /// Tests EVALUATE with TOPN for limited results.
    /// LLM use case: "show me top 5 sales by amount"
    /// </summary>
    [Fact]
    public void Evaluate_TopNQuery_ReturnsLimitedRows()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE TOPN(3, 'SalesTable', 'SalesTable'[Amount], DESC)");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);
        Assert.True(result.RowCount <= 3, $"Expected at most 3 rows, got {result.RowCount}");
    }

    /// <summary>
    /// Tests EVALUATE with DISTINCT for unique values.
    /// LLM use case: "show me unique customer IDs"
    /// </summary>
    [Fact]
    public void Evaluate_DistinctQuery_ReturnsUniqueValues()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Evaluate(batch,
            "EVALUATE DISTINCT('SalesTable'[CustomerID])");

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.NotNull(result.Rows);
        Assert.Equal(1, result.ColumnCount); // DISTINCT on single column returns single column

        // Verify all values are unique
        var values = result.Rows.Select(r => r[0]).ToList();
        var uniqueValues = values.Distinct().ToList();
        Assert.Equal(values.Count, uniqueValues.Count);
    }

    #endregion
}





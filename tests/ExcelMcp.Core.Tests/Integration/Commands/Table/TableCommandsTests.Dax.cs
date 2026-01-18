using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Integration tests for DAX-backed Table operations (create-from-dax, update-dax, get-dax).
/// Tests verify that Excel Tables can be created from DAX EVALUATE queries and their queries updated.
/// Uses DataModelPivotTableFixture which provides Data Model tables for DAX queries.
/// </summary>
[Collection("DataModel")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Slow")]
public class TableCommandsTests_Dax
{
    private readonly TableCommands _tableCommands;
    private readonly string _dataModelFile;

    public TableCommandsTests_Dax(DataModelPivotTableFixture fixture)
    {
        _tableCommands = new TableCommands();
        _dataModelFile = fixture.TestFilePath;
    }

    #region CreateFromDax Tests

    /// <summary>
    /// Tests creating a DAX-backed table with a simple EVALUATE query.
    /// LLM use case: "create a table from this DAX query"
    /// </summary>
    [Fact]
    public void CreateFromDax_SimpleEvaluateQuery_CreatesTable()
    {
        var tableName = $"DaxTable_{Guid.NewGuid():N}";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create a DAX-backed table
        _tableCommands.CreateFromDax(
            batch,
            "Sheet1", // Use existing sheet or create if needed
            tableName,
            "EVALUATE 'SalesTable'",
            "A1");

        // Verify table was created
        var listResult = _tableCommands.List(batch);
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        Assert.Contains(listResult.Tables, t => t.Name == tableName);

        // Verify GetDax shows it's a DAX-backed table
        var daxInfo = _tableCommands.GetDax(batch, tableName);
        Assert.True(daxInfo.Success, $"GetDax failed: {daxInfo.ErrorMessage}");
        Assert.True(daxInfo.HasDaxConnection, "Expected table to have DAX connection");
        Assert.NotNull(daxInfo.DaxQuery);
        Assert.NotEmpty(daxInfo.DaxQuery!);
    }

    /// <summary>
    /// Tests creating a DAX-backed table with SUMMARIZE aggregation.
    /// LLM use case: "create a summary table with totals by customer"
    /// </summary>
    [Fact]
    public void CreateFromDax_SummarizeQuery_CreatesAggregatedTable()
    {
        var tableName = $"SummaryTable_{Guid.NewGuid():N}";
        var daxQuery = "EVALUATE SUMMARIZE('SalesTable', 'SalesTable'[CustomerID], \"TotalAmount\", SUM('SalesTable'[Amount]))";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create a DAX-backed table with SUMMARIZE
        _tableCommands.CreateFromDax(
            batch,
            "Sheet1",
            tableName,
            daxQuery,
            "A1");

        // Verify table was created
        var readResult = _tableCommands.Read(batch, tableName);
        Assert.True(readResult.Success, $"Read failed: {readResult.ErrorMessage}");
        Assert.NotNull(readResult.Table);

        // GetDax should return the query
        var daxInfo = _tableCommands.GetDax(batch, tableName);
        Assert.True(daxInfo.HasDaxConnection);
    }

    /// <summary>
    /// Tests creating a DAX-backed table with FILTER.
    /// LLM use case: "create a filtered view of the data"
    /// </summary>
    [Fact]
    public void CreateFromDax_FilterQuery_CreatesFilteredTable()
    {
        var tableName = $"FilteredTable_{Guid.NewGuid():N}";
        var daxQuery = "EVALUATE FILTER('SalesTable', 'SalesTable'[Amount] > 100)";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        _tableCommands.CreateFromDax(
            batch,
            "Sheet1",
            tableName,
            daxQuery,
            "A1");

        var listResult = _tableCommands.List(batch);
        Assert.Contains(listResult.Tables, t => t.Name == tableName);
    }

    /// <summary>
    /// Tests creating DAX table with custom target cell.
    /// </summary>
    [Fact]
    public void CreateFromDax_CustomTargetCell_PlacesTableCorrectly()
    {
        var tableName = $"OffsetTable_{Guid.NewGuid():N}";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create table starting at C5
        _tableCommands.CreateFromDax(
            batch,
            "Sheet1",
            tableName,
            "EVALUATE 'CustomersTable'",
            "C5");

        // Verify table exists
        var listResult = _tableCommands.List(batch);
        Assert.Contains(listResult.Tables, t => t.Name == tableName);

        // Verify table position (Read should show the range)
        var readResult = _tableCommands.Read(batch, tableName);
        Assert.True(readResult.Success);
        Assert.NotNull(readResult.Table?.Range);
        // Range should include C5
        Assert.Contains("C", readResult.Table!.Range, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region UpdateDax Tests

    /// <summary>
    /// Tests updating the DAX query of an existing DAX-backed table.
    /// LLM use case: "change the filter on this DAX table"
    /// </summary>
    [Fact]
    public void UpdateDax_ExistingDaxTable_UpdatesQuery()
    {
        var tableName = $"UpdateDaxTable_{Guid.NewGuid():N}";
        var originalQuery = "EVALUATE 'SalesTable'";
        var updatedQuery = "EVALUATE FILTER('SalesTable', 'SalesTable'[CustomerID] = 1)";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create initial DAX table
        _tableCommands.CreateFromDax(batch, "Sheet1", tableName, originalQuery, "A1");

        // Update the DAX query
        _tableCommands.UpdateDax(batch, tableName, updatedQuery);

        // Verify the query was updated
        var daxInfo = _tableCommands.GetDax(batch, tableName);
        Assert.True(daxInfo.Success);
        Assert.Contains("FILTER", daxInfo.DaxQuery, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that UpdateDax fails for non-DAX tables.
    /// </summary>
    [Fact]
    public void UpdateDax_NonDaxTable_ThrowsError()
    {
        // SalesTable from fixture is a regular table, not DAX-backed
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Attempting to update a non-DAX table should throw some kind of exception
        // (could be InvalidOperationException from our validation or COMException from COM)
        var ex = Assert.ThrowsAny<Exception>(() =>
            _tableCommands.UpdateDax(batch, "SalesTable", "EVALUATE 'ProductsTable'"));

        Assert.NotNull(ex);
    }

    /// <summary>
    /// Tests UpdateDax with invalid DAX syntax.
    /// </summary>
    [Fact]
    public void UpdateDax_InvalidDax_ThrowsError()
    {
        var tableName = $"UpdateErrorTable_{Guid.NewGuid():N}";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create initial DAX table
        _tableCommands.CreateFromDax(batch, "Sheet1", tableName, "EVALUATE 'SalesTable'", "A1");

        // Try to update with invalid DAX - should throw
        var ex = Assert.ThrowsAny<Exception>(() =>
            _tableCommands.UpdateDax(batch, tableName, "EVALUATE INVALID_SYNTAX()"));

        Assert.NotNull(ex);
    }

    #endregion

    #region GetDax Tests

    /// <summary>
    /// Tests GetDax on a DAX-backed table returns query info.
    /// LLM use case: "what DAX query is this table using?"
    /// </summary>
    [Fact]
    public void GetDax_DaxBackedTable_ReturnsQueryInfo()
    {
        var tableName = $"GetDaxTable_{Guid.NewGuid():N}";
        var daxQuery = "EVALUATE TOPN(10, 'SalesTable', 'SalesTable'[Amount], DESC)";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create DAX table
        _tableCommands.CreateFromDax(batch, "Sheet1", tableName, daxQuery, "A1");

        // Get DAX info
        var result = _tableCommands.GetDax(batch, tableName);

        Assert.True(result.Success, $"GetDax failed: {result.ErrorMessage}");
        Assert.Equal(tableName, result.TableName);
        Assert.True(result.HasDaxConnection);
        Assert.NotNull(result.DaxQuery);
        Assert.Contains("TOPN", result.DaxQuery!, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(result.ModelConnectionName);
        Assert.NotEmpty(result.ModelConnectionName!);
    }

    /// <summary>
    /// Tests GetDax on a regular (non-DAX) table returns HasDaxConnection = false.
    /// LLM use case: "check if this table is DAX-backed"
    /// </summary>
    [Fact]
    public void GetDax_RegularTable_ReturnsNoDaxConnection()
    {
        // SalesTable from fixture is a regular table loaded to Data Model
        // but it's not backed by a DAX query
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var result = _tableCommands.GetDax(batch, "SalesTable");

        Assert.True(result.Success);
        Assert.Equal("SalesTable", result.TableName);
        Assert.False(result.HasDaxConnection);
        Assert.True(string.IsNullOrEmpty(result.DaxQuery));
    }

    /// <summary>
    /// Tests GetDax on non-existent table throws error.
    /// </summary>
    [Fact]
    public void GetDax_NonExistentTable_ThrowsError()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.ThrowsAny<Exception>(() =>
            _tableCommands.GetDax(batch, "NonExistentTable_12345"));

        Assert.NotNull(ex);
    }

    #endregion

    #region Parameter Validation Tests

    /// <summary>
    /// Tests CreateFromDax with null sheetName throws ArgumentException.
    /// </summary>
    [Fact]
    public void CreateFromDax_NullSheetName_ThrowsArgumentException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.Throws<ArgumentException>(() =>
            _tableCommands.CreateFromDax(batch, null!, "TestTable", "EVALUATE 'Sales'"));

        Assert.Contains("sheetName", ex.Message);
    }

    /// <summary>
    /// Tests CreateFromDax with null tableName throws ArgumentException.
    /// </summary>
    [Fact]
    public void CreateFromDax_NullTableName_ThrowsArgumentException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.Throws<ArgumentException>(() =>
            _tableCommands.CreateFromDax(batch, "Sheet1", null!, "EVALUATE 'Sales'"));

        Assert.Contains("tableName", ex.Message);
    }

    /// <summary>
    /// Tests CreateFromDax with null daxQuery throws ArgumentException.
    /// </summary>
    [Fact]
    public void CreateFromDax_NullDaxQuery_ThrowsArgumentException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.Throws<ArgumentException>(() =>
            _tableCommands.CreateFromDax(batch, "Sheet1", "TestTable", null!));

        Assert.Contains("daxQuery", ex.Message);
    }

    /// <summary>
    /// Tests UpdateDax with null daxQuery throws ArgumentException.
    /// </summary>
    [Fact]
    public void UpdateDax_NullDaxQuery_ThrowsArgumentException()
    {
        var tableName = $"UpdateNullTable_{Guid.NewGuid():N}";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create table first
        _tableCommands.CreateFromDax(batch, "Sheet1", tableName, "EVALUATE 'SalesTable'", "A1");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tableCommands.UpdateDax(batch, tableName, null!));

        Assert.Contains("daxQuery", ex.Message);
    }

    #endregion
}

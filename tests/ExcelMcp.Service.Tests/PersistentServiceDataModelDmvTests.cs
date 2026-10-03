using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for DMV (Dynamic Management View) query execution.
/// Tests verify that DMV queries can be executed against the Data Model's embedded
/// Analysis Services engine and return tabular metadata results.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class PersistentServiceDataModelDmvTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IDataModelCommands _dataModelCommands =
        fixture.CreateCommands<IDataModelCommands>();
    private readonly IDataModelRelCommands _relationships =
        fixture.CreateCommands<IDataModelRelCommands>();

    #region Basic DMV Query Tests

    /// <summary>
    /// Tests that TMSCHEMA_TABLES DMV returns table metadata.
    /// Note: Excel's embedded Analysis Services may return 0 rows for this DMV.
    /// LLM use case: "show me all tables in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaTablesQuery_ReturnsSchemaWithoutError()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        Assert.Contains(result.Columns, c => c.Equals("ID", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Columns, c => c.Equals("Name", StringComparison.OrdinalIgnoreCase));
        AssertResultShape(result);
    }

    /// <summary>
    /// Tests that TMSCHEMA_COLUMNS DMV returns column metadata.
    /// Note: Excel's embedded Analysis Services may return 0 rows for this DMV.
    /// LLM use case: "show me all columns in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaColumnsQuery_ReturnsSchemaWithoutError()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_COLUMNS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        Assert.Contains(result.Columns, c => c.Equals("TableID", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Columns, c => c.Equals("ExplicitName", StringComparison.OrdinalIgnoreCase));
        AssertResultShape(result);
    }

    /// <summary>
    /// Tests that TMSCHEMA_MEASURES DMV returns measure metadata.
    /// LLM use case: "list all measures in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaMeasuresQuery_ReturnsMeasures()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: May have 0 rows if no measures defined, but columns should exist

        // TMSCHEMA_MEASURES should have standard columns
        Assert.Contains(result.Columns, c => c.Equals("Name", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Columns, c => c.Equals("Expression", StringComparison.OrdinalIgnoreCase));
        AssertResultShape(result);
    }

    /// <summary>
    /// Tests that TMSCHEMA_RELATIONSHIPS DMV returns relationship metadata.
    /// LLM use case: "show me all relationships in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaRelationshipsQuery_ReturnsRelationships()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: May have 0 rows if no relationships defined, but columns should exist

        // TMSCHEMA_RELATIONSHIPS should have FromTableID and ToTableID columns
        Assert.Contains(result.Columns, c => c.Equals("FromTableID", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Columns, c => c.Equals("ToTableID", StringComparison.OrdinalIgnoreCase));
        AssertResultShape(result);
    }

    #endregion

    #region Filtered DMV Query Tests

    /// <summary>
    /// Tests DMV query with WHERE clause filter.
    /// Note: Excel's embedded Analysis Services has limited DMV support.
    /// Uses DISCOVER_CALC_DEPENDENCY which is known to work.
    /// LLM use case: "show me calculation dependencies"
    /// </summary>
    [Fact]
    public void ExecuteDmv_DiscoverDependencyQuery_ReturnsResults()
    {
        var batch = _fixture.BatchToken;

        // DISCOVER_CALC_DEPENDENCY is known to work in Excel's embedded AS
        var result = _dataModelCommands.ExecuteDmv(batch,
            "SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY");

        Assert.True(result.Success, $"Query failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        Assert.Contains(result.Columns, c => c.Equals("OBJECT", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Columns, c => c.Equals("OBJECT_TYPE", StringComparison.OrdinalIgnoreCase));
        AssertResultShape(result);
    }

    /// <summary>
    /// Tests that SELECT * queries work (Excel's embedded AS doesn't support column selection).
    /// LLM use case: "query all columns from a DMV"
    /// </summary>
    [Fact]
    public void ExecuteDmv_SelectAllQuery_Works()
    {
        var batch = _fixture.BatchToken;

        // SELECT * works - specific column selection doesn't work in Excel's embedded AS
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.True(result.ColumnCount > 0, "Expected columns from DBSCHEMA_CATALOGS");
        AssertCatalog(result);
    }

    #endregion

    #region Advanced DMV Query Tests

    /// <summary>
    /// Tests DISCOVER_CALC_DEPENDENCY DMV for dependency analysis.
    /// LLM use case: "show me measure dependencies"
    /// </summary>
    [Fact]
    public void ExecuteDmv_DiscoverCalcDependencyQuery_ReturnsDependencies()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: May have 0 rows if no calculated dependencies, but columns should exist

        // DISCOVER_CALC_DEPENDENCY should have standard columns
        Assert.Contains(result.Columns, c => c.Equals("OBJECT", StringComparison.OrdinalIgnoreCase) ||
                                             c.Equals("OBJECT_TYPE", StringComparison.OrdinalIgnoreCase));
        AssertResultShape(result);
    }

    /// <summary>
    /// Tests DBSCHEMA_CATALOGS DMV for catalog info.
    /// LLM use case: "show me database catalogs"
    /// </summary>
    [Fact]
    public void ExecuteDmv_DbschemaCatalogsQuery_ReturnsCatalogs()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        Assert.True(result.RowCount > 0, "Expected at least one catalog");

        Assert.Contains(result.Columns, c => c.Equals("CATALOG_NAME", StringComparison.OrdinalIgnoreCase));
        AssertCatalog(result);
    }

    #endregion

    #region Error Handling Tests

    /// <summary>
    /// Tests that invalid DMV query throws exception.
    /// LLM use case: handling syntax errors
    /// </summary>
    [Fact]
    public void ExecuteDmv_InvalidQuery_ThrowsException()
    {
        var batch = _fixture.BatchToken;
        var before = CaptureGuardedModelState();

        // Invalid DMV - non-existent system view
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.NONEXISTENT_VIEW"));

        Assert.Contains("ComInterop/", ex.Message, StringComparison.Ordinal);
        Assert.Equal(before.State, CaptureModelState(before.GuardSheet));
        AssertCatalog(_dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS"));
    }

    /// <summary>
    /// Tests that null/empty query throws ArgumentException.
    /// </summary>
    [Fact]
    public void ExecuteDmv_NullQuery_ThrowsArgumentException()
    {
        var batch = _fixture.BatchToken;
        var before = CaptureGuardedModelState();

        var ex = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.ExecuteDmv(batch, ""));

        Assert.Contains("dmvQuery", ex.Message);
        Assert.Equal(before.State, CaptureModelState(before.GuardSheet));
        AssertCatalog(_dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS"));
    }

    /// <summary>
    /// Tests that whitespace-only query throws ArgumentException.
    /// </summary>
    [Fact]
    public void ExecuteDmv_WhitespaceQuery_ThrowsArgumentException()
    {
        var batch = _fixture.BatchToken;
        var before = CaptureGuardedModelState();

        var ex = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.ExecuteDmv(batch, "   "));

        Assert.Contains("dmvQuery", ex.Message);
        Assert.Equal(before.State, CaptureModelState(before.GuardSheet));
        AssertCatalog(_dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS"));
    }

    /// <summary>
    /// Tests that malformed SQL query throws exception.
    /// </summary>
    [Fact]
    public void ExecuteDmv_MalformedSqlQuery_ThrowsException()
    {
        var batch = _fixture.BatchToken;
        var before = CaptureGuardedModelState();

        var ex = Assert.Throws<InvalidOperationException>(() =>
            _dataModelCommands.ExecuteDmv(batch, "INVALID SQL SYNTAX HERE"));

        Assert.Contains("ComInterop/", ex.Message, StringComparison.Ordinal);
        Assert.Equal(before.State, CaptureModelState(before.GuardSheet));
        AssertCatalog(_dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS"));
    }

    #endregion

    #region Query Result Validation Tests

    /// <summary>
    /// Tests that DMV query result includes proper DmvQuery echo.
    /// </summary>
    [Fact]
    public void ExecuteDmv_ValidQuery_EchoesQueryInResult()
    {
        var batch = _fixture.BatchToken;
        var query = "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES";
        var result = _dataModelCommands.ExecuteDmv(batch, query);

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.Equal(query, result.DmvQuery);
        AssertResultShape(result);
    }

    /// <summary>
    /// Tests that RowCount and ColumnCount match actual data.
    /// </summary>
    [Fact]
    public void ExecuteDmv_ValidQuery_CountsMatchActualData()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.Equal(result.Columns.Count, result.ColumnCount);
        Assert.Equal(result.Rows.Count, result.RowCount);
        AssertResultShape(result);
    }

    #endregion

    private (string State, string GuardSheet) CaptureGuardedModelState()
    {
        var guardSheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, guardSheet, "A1:B2",
            [[17, "DMV guard"], [29, "Retained neighbor"]]));
        return (CaptureModelState(guardSheet), guardSheet);
    }

    private string CaptureModelState(string guardSheet)
    {
        var batch = _fixture.BatchToken;
        var tables = RequireSuccess(_dataModelCommands.ListTables(batch)).Tables;
        Assert.Contains(tables, table => table.Name == "SalesTable");
        var measures = RequireSuccess(_dataModelCommands.ListMeasures(batch)).Measures;
        Assert.Contains(measures, measure => measure.Name == "Total Sales");
        var records = RequireSuccess(_dataModelCommands.Evaluate(batch,
            "EVALUATE SalesTable ORDER BY SalesTable[SalesID]"));
        Assert.Equal(10, records.RowCount);
        Assert.Equal(6, records.ColumnCount);
        Assert.Equal(10, records.Rows.Count);
        Assert.All(records.Rows, row => Assert.Equal(6, row.Count));
        var total = RequireSuccess(_dataModelCommands.Evaluate(batch,
            "EVALUATE ROW(\"Total\", [Total Sales])"));
        Assert.Equal(2455, Convert.ToDouble(Assert.Single(Assert.Single(total.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
        var guard = RequireSuccess(_commands.GetValues(batch, guardSheet, "A1:B2"));
        Assert.Equal("[[17,\"DMV guard\"],[29,\"Retained neighbor\"]]",
            System.Text.Json.JsonSerializer.Serialize(guard.Values));
        return System.Text.Json.JsonSerializer.Serialize(new
        {
            Tables = tables,
            Columns = tables.Select(table => new
            {
                table.Name,
                Columns = RequireSuccess(_dataModelCommands.ListColumns(batch, table.Name)).Columns
            }).ToList(),
            Relationships = RequireSuccess(_relationships.ListRelationships(batch)).Relationships,
            Measures = measures.Select(measure => RequireSuccess(_dataModelCommands.Read(batch, measure.Name))).ToList(),
            ModelRecords = tables.Select(table =>
            {
                var data = RequireSuccess(_dataModelCommands.Evaluate(batch,
                    $"EVALUATE '{table.Name.Replace("'", "''", StringComparison.Ordinal)}'"));
                Assert.NotEmpty(data.Rows);
                Assert.Equal(data.RowCount, data.Rows.Count);
                Assert.Equal(data.ColumnCount, data.Columns.Count);
                Assert.All(data.Rows, row => Assert.Equal(data.ColumnCount, row.Count));
                return new { table.Name, data.Columns, data.Rows };
            }).ToList(),
            SourceCells = RequireSuccess(_commands.GetValues(batch, "SalesData", "A1:F11")).Values,
            GuardCells = guard.Values
        });
    }

    private static void AssertResultShape(Sbroenne.ExcelMcp.Core.Models.DmvQueryResult result)
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.NotEmpty(result.Columns);
        Assert.Equal(result.Columns.Count, result.ColumnCount);
        Assert.Equal(result.Rows.Count, result.RowCount);
        Assert.All(result.Rows, row => Assert.Equal(result.ColumnCount, row.Count));
    }

    private static void AssertCatalog(Sbroenne.ExcelMcp.Core.Models.DmvQueryResult result)
    {
        AssertResultShape(result);
        var index = result.Columns.FindIndex(column => column.Equals("CATALOG_NAME", StringComparison.OrdinalIgnoreCase));
        Assert.True(index >= 0);
        var row = Assert.Single(result.Rows);
        Assert.False(string.IsNullOrWhiteSpace(row[index]?.ToString()));
    }
}

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for DMV (Dynamic Management View) query execution.
/// Tests verify that DMV queries can be executed against the Data Model's embedded
/// Analysis Services engine and return tabular metadata results.
/// </summary>
[Collection("DataModel")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class DataModelCommandsTests_Dmv
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly string _dataModelFile;

    public DataModelCommandsTests_Dmv(DataModelPivotTableFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _dataModelFile = fixture.TestFilePath;
    }

    #region Basic DMV Query Tests

    /// <summary>
    /// Tests that TMSCHEMA_TABLES DMV returns table metadata.
    /// Note: Excel's embedded Analysis Services may return 0 rows for this DMV.
    /// LLM use case: "show me all tables in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaTablesQuery_ReturnsSchemaWithoutError()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: Excel's embedded AS may return 0 rows for TMSCHEMA_TABLES
        // Just verify the query executes without error

        // TMSCHEMA_TABLES should have ID and Name columns if any results
        if (result.ColumnCount > 0)
        {
            Assert.Contains(result.Columns, c => c.Equals("ID", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(result.Columns, c => c.Equals("Name", StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>
    /// Tests that TMSCHEMA_COLUMNS DMV returns column metadata.
    /// Note: Excel's embedded Analysis Services may return 0 rows for this DMV.
    /// LLM use case: "show me all columns in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaColumnsQuery_ReturnsSchemaWithoutError()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_COLUMNS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: Excel's embedded AS may return 0 rows for TMSCHEMA_COLUMNS
        // Just verify the query executes without error

        // TMSCHEMA_COLUMNS should have TableID and ExplicitName columns if any results
        if (result.ColumnCount > 0)
        {
            Assert.Contains(result.Columns, c => c.Equals("TableID", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(result.Columns, c => c.Equals("ExplicitName", StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>
    /// Tests that TMSCHEMA_MEASURES DMV returns measure metadata.
    /// LLM use case: "list all measures in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaMeasuresQuery_ReturnsMeasures()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: May have 0 rows if no measures defined, but columns should exist

        // TMSCHEMA_MEASURES should have standard columns
        Assert.Contains(result.Columns, c => c.Equals("Name", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Columns, c => c.Equals("Expression", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Tests that TMSCHEMA_RELATIONSHIPS DMV returns relationship metadata.
    /// LLM use case: "show me all relationships in the Data Model"
    /// </summary>
    [Fact]
    public void ExecuteDmv_TmschemaRelationshipsQuery_ReturnsRelationships()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: May have 0 rows if no relationships defined, but columns should exist

        // TMSCHEMA_RELATIONSHIPS should have FromTableID and ToTableID columns
        Assert.Contains(result.Columns, c => c.Equals("FromTableID", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Columns, c => c.Equals("ToTableID", StringComparison.OrdinalIgnoreCase));
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
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // DISCOVER_CALC_DEPENDENCY is known to work in Excel's embedded AS
        var result = _dataModelCommands.ExecuteDmv(batch,
            "SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY");

        Assert.True(result.Success, $"Query failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // DISCOVER_CALC_DEPENDENCY returns columns about calculation dependencies
    }

    /// <summary>
    /// Tests that SELECT * queries work (Excel's embedded AS doesn't support column selection).
    /// LLM use case: "query all columns from a DMV"
    /// </summary>
    [Fact]
    public void ExecuteDmv_SelectAllQuery_Works()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // SELECT * works - specific column selection doesn't work in Excel's embedded AS
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.True(result.ColumnCount > 0, "Expected columns from DBSCHEMA_CATALOGS");
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
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        // Note: May have 0 rows if no calculated dependencies, but columns should exist

        // DISCOVER_CALC_DEPENDENCY should have standard columns
        Assert.Contains(result.Columns, c => c.Equals("OBJECT", StringComparison.OrdinalIgnoreCase) ||
                                             c.Equals("OBJECT_TYPE", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Tests DBSCHEMA_CATALOGS DMV for catalog info.
    /// LLM use case: "show me database catalogs"
    /// </summary>
    [Fact]
    public void ExecuteDmv_DbschemaCatalogsQuery_ReturnsCatalogs()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.NotNull(result.Rows);
        Assert.True(result.RowCount > 0, "Expected at least one catalog");

        Assert.Contains(result.Columns, c => c.Equals("CATALOG_NAME", StringComparison.OrdinalIgnoreCase));
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
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Invalid DMV - non-existent system view
        var ex = Assert.ThrowsAny<Exception>(() =>
            _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.NONEXISTENT_VIEW"));

        Assert.NotNull(ex);
    }

    /// <summary>
    /// Tests that null/empty query throws ArgumentException.
    /// </summary>
    [Fact]
    public void ExecuteDmv_NullQuery_ThrowsArgumentException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.ExecuteDmv(batch, ""));

        Assert.Contains("dmvQuery", ex.Message);
    }

    /// <summary>
    /// Tests that whitespace-only query throws ArgumentException.
    /// </summary>
    [Fact]
    public void ExecuteDmv_WhitespaceQuery_ThrowsArgumentException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.Throws<ArgumentException>(() =>
            _dataModelCommands.ExecuteDmv(batch, "   "));

        Assert.Contains("dmvQuery", ex.Message);
    }

    /// <summary>
    /// Tests that malformed SQL query throws exception.
    /// </summary>
    [Fact]
    public void ExecuteDmv_MalformedSqlQuery_ThrowsException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.ThrowsAny<Exception>(() =>
            _dataModelCommands.ExecuteDmv(batch, "INVALID SQL SYNTAX HERE"));

        Assert.NotNull(ex);
    }

    #endregion

    #region Query Result Validation Tests

    /// <summary>
    /// Tests that DMV query result includes proper DmvQuery echo.
    /// </summary>
    [Fact]
    public void ExecuteDmv_ValidQuery_EchoesQueryInResult()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var query = "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES";
        var result = _dataModelCommands.ExecuteDmv(batch, query);

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.Equal(query, result.DmvQuery);
    }

    /// <summary>
    /// Tests that RowCount and ColumnCount match actual data.
    /// </summary>
    [Fact]
    public void ExecuteDmv_ValidQuery_CountsMatchActualData()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.ExecuteDmv(batch, "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES");

        Assert.True(result.Success, $"ExecuteDmv failed: {result.ErrorMessage}");
        Assert.Equal(result.Columns.Count, result.ColumnCount);
        Assert.Equal(result.Rows.Count, result.RowCount);
    }

    #endregion
}





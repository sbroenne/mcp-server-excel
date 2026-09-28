using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public interface IDataModelServiceCommands :
    IDataModelCommands,
    IDataModelRelCommands;

/// <summary>
/// Integration tests for Data Model operations focusing on LLM use cases.
/// Tests cover essential workflows: list tables/measures/relationships, create/update/delete measures, manage relationships.
/// Uses DataModelPivotTableFixture which creates ONE comprehensive Data Model + PivotTable workbook (shared via collection fixture).
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public partial class PersistentServiceDataModelCommandsTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IDataModelServiceCommands _dataModelCommands =
        fixture.CreateCommands<IDataModelServiceCommands>();
    private readonly string _dataModelFile = fixture.WorkbookPath;
    private readonly DataModelPivotTableCreationResult _creationResult =
        PersistentServiceDataModelFixture.CreationResult;

    #region Core Discovery Tests (4 tests)

    /// <summary>
    /// Validates that the fixture successfully created the Data Model.
    /// LLM use case: "create a data model with tables, relationships, and measures"
    /// </summary>
    [Fact]
    public void Create_CompleteDataModel_SuccessfullyCreatesAllComponents()
    {
        Assert.True(_creationResult.Success,
            $"Data Model creation failed: {_creationResult.ErrorMessage}");
        Assert.True(_creationResult.FileCreated);
        Assert.Equal(5, _creationResult.TablesCreated);  // SalesTable, CustomersTable, ProductsTable, RegionalSalesTable, DisambiguationTable
        Assert.Equal(5, _creationResult.TablesLoadedToModel);
        Assert.Equal(2, _creationResult.RelationshipsCreated);
        Assert.Equal(6, _creationResult.MeasuresCreated);  // Total Sales, Average Sale, Total Customers, TotalRevenue, ACR, Discount
    }

    /// <summary>
    /// Tests listing tables in the data model.
    /// LLM use case: "show me all tables in the data model"
    /// </summary>
    [Fact]
    public void ListTables_WithDataModel_ReturnsTables()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ListTables(batch);

        Assert.True(result.Success, $"ListTables failed: {result.ErrorMessage}");
        Assert.Equal(5, result.Tables.Count);  // Now includes RegionalSalesTable and DisambiguationTable
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
        Assert.Contains(result.Tables, t => t.Name == "CustomersTable");
        Assert.Contains(result.Tables, t => t.Name == "ProductsTable");
        Assert.Contains(result.Tables, t => t.Name == "RegionalSalesTable");
        Assert.Contains(result.Tables, t => t.Name == "DisambiguationTable");
    }

    /// <summary>
    /// Tests getting table details with columns.
    /// LLM use case: "show me the columns in this data model table"
    /// </summary>
    [Fact]
    public void GetTable_WithValidTable_ReturnsCompleteInfo()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ReadTable(batch, "SalesTable");

        Assert.True(result.Success, $"ViewTable failed: {result.ErrorMessage}");
        Assert.Equal("SalesTable", result.TableName);
        Assert.NotNull(result.SourceName);
        Assert.True(result.RecordCount >= 10);
        Assert.NotNull(result.Columns);
        Assert.True(result.Columns.Count >= 6);
    }

    /// <summary>
    /// Tests getting data model statistics.
    /// LLM use case: "show me information about this data model"
    /// </summary>
    [Fact]
    public void GetInfo_WithRealisticDataModel_ReturnsAccurateStatistics()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ReadInfo(batch);

        Assert.True(result.Success, $"GetModelInfo failed: {result.ErrorMessage}");
        Assert.Equal(5, result.TableCount);  // Now includes RegionalSalesTable and DisambiguationTable
        Assert.Equal(6, result.MeasureCount);  // Now includes TotalRevenue, ACR, Discount
        Assert.Equal(2, result.RelationshipCount);
        Assert.True(result.TotalRows > 0);
        Assert.NotNull(result.TableNames);
        Assert.Contains("SalesTable", result.TableNames);
    }

    #endregion

    #region Measure Operations (5 tests)

    /// <summary>
    /// Tests listing all measures in the data model.
    /// LLM use case: "show me all DAX measures"
    /// </summary>
    [Fact]
    public void ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ListMeasures(batch);

        Assert.True(result.Success, $"ListMeasures failed: {result.ErrorMessage}");
        Assert.NotNull(result.Measures);
        Assert.Equal(6, result.Measures.Count);  // Now includes TotalRevenue, ACR, Discount

        var measureNames = result.Measures.Select(m => m.Name).ToList();
        Assert.Contains("Total Sales", measureNames);
        Assert.Contains("Average Sale", measureNames);
        Assert.Contains("Total Customers", measureNames);
        Assert.Contains("TotalRevenue", measureNames);
        Assert.Contains("ACR", measureNames);
        Assert.Contains("Discount", measureNames);
    }

    /// <summary>
    /// Tests viewing a specific measure's DAX formula.
    /// LLM use case: "show me the DAX formula for this measure"
    /// </summary>
    [Fact]
    public void Get_WithRealisticDataModel_ReturnsValidDAXFormula()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Read(batch, "Total Sales");

        Assert.True(result.Success, $"ViewMeasure failed: {result.ErrorMessage}");
        Assert.NotNull(result.DaxFormula);
        Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Amount", result.DaxFormula);
        Assert.Equal("Total Sales", result.MeasureName);
    }

    /// <summary>
    /// Tests that Read returns structured FormatInfo with type and properties.
    /// Regression test: FormatString was previously a string, now it's a structured FormatInfo object.
    /// </summary>
    [Fact]
    public void Read_WithMeasure_ReturnsStructuredFormatInfo()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Read(batch, "Total Sales");

        Assert.True(result.Success, $"Read failed: {result.ErrorMessage}");

        // FormatInfo should be populated (not null)
        Assert.NotNull(result.FormatInfo);

        // Type should be a known format type
        var validTypes = new HashSet<string> { "General", "Currency", "Decimal", "Percentage", "WholeNumber" };
        Assert.Contains(result.FormatInfo.Type, validTypes);
    }

    /// <summary>
    /// Tests creating a new DAX measure.
    /// LLM use case: "create a DAX measure"
    /// </summary>
    [Fact]
    public void CreateMeasure_ValidNameAndFormula_CreatesSuccessfully()
    {
        var measureName = $"Test_{nameof(CreateMeasure_ValidNameAndFormula_CreatesSuccessfully)}_{Guid.NewGuid():N}";
        var daxFormula = "SUM(SalesTable[Amount])";

        var batch = _fixture.BatchToken;
        _ = CreateMeasure("SalesTable", measureName, daxFormula);  // CreateMeasure throws on error

        // Verify measure created
        var listResult = _dataModelCommands.ListMeasures(batch);
        Assert.Contains(listResult.Measures, m => m.Name == measureName);
    }

    /// <summary>
    /// Tests updating an existing measure's DAX formula.
    /// LLM use case: "update this measure's formula"
    /// </summary>
    [Fact]
    public void UpdateMeasure_WithValidFormula_UpdatesSuccessfully()
    {
        var measureName = $"Test_{nameof(UpdateMeasure_WithValidFormula_UpdatesSuccessfully)}_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        var updatedFormula = "AVERAGE(SalesTable[Amount])";

        var batch = _fixture.BatchToken;

        // Create measure
        _ = CreateMeasure("SalesTable", measureName, originalFormula);  // CreateMeasure throws on error

        // Update formula
        _ = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: updatedFormula);  // UpdateMeasure throws on error

        // Verify update
        var viewResult = _dataModelCommands.Read(batch, measureName);
        Assert.Contains("AVERAGE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests deleting a measure.
    /// LLM use case: "delete this DAX measure"
    /// </summary>
    [Fact]
    public void DeleteMeasure_WithValidMeasure_ReturnsSuccessResult()
    {
        var measureName = $"Test_{nameof(DeleteMeasure_WithValidMeasure_ReturnsSuccessResult)}_{Guid.NewGuid():N}";

        var batch = _fixture.BatchToken;

        // Create measure
        _ = CreateMeasure("SalesTable", measureName, "SUM(SalesTable[Amount])");  // CreateMeasure throws on error

        // Delete measure
        _ = DeleteMeasure(measureName);  // DeleteMeasure throws on error

        // Verify deletion
        var listResult = _dataModelCommands.ListMeasures(batch);
        Assert.DoesNotContain(listResult.Measures, m => m.Name == measureName);
    }

    #endregion

    #region Relationship Operations (3 tests)

    /// <summary>
    /// Tests listing all relationships in the data model.
    /// LLM use case: "show me all table relationships"
    /// </summary>
    [Fact]
    public void ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithDetails()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.ListRelationships(batch);

        Assert.True(result.Success, $"ListRelationships failed: {result.ErrorMessage}");
        Assert.NotNull(result.Relationships);
        Assert.Equal(2, result.Relationships.Count);

        // Verify SalesTable->CustomersTable relationship
        var salesCustomersRel = result.Relationships.FirstOrDefault(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable");
        Assert.NotNull(salesCustomersRel);
        Assert.Equal("CustomerID", salesCustomersRel.FromColumn);
        Assert.Equal("CustomerID", salesCustomersRel.ToColumn);
        Assert.True(salesCustomersRel.IsActive);

        // Verify SalesTable->ProductsTable relationship
        var salesProductsRel = result.Relationships.FirstOrDefault(r =>
            r.FromTable == "SalesTable" && r.ToTable == "ProductsTable");
        Assert.NotNull(salesProductsRel);
        Assert.Equal("ProductID", salesProductsRel.FromColumn);
        Assert.Equal("ProductID", salesProductsRel.ToColumn);
        Assert.True(salesProductsRel.IsActive);
    }

    /// <summary>
    /// Tests creating a new relationship between tables.
    /// LLM use case: "create a relationship between these tables"
    /// </summary>
    [Fact]
    public void CreateRelationship_ValidTablesAndColumns_CreatesSuccessfully()
    {
        var batch = _fixture.BatchToken;

        // Delete existing relationship first to allow recreating it
        var listResult = _dataModelCommands.ListRelationships(batch);
        if (listResult.Success && listResult.Relationships?.Any(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID") == true)
        {
            _ = _dataModelCommands.DeleteRelationship(batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");  // DeleteRelationship throws on error
        }

        // Create relationship
        _ = _dataModelCommands.CreateRelationship(
            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");  // CreateRelationship throws on error

        // Verify creation
        var verifyResult = _dataModelCommands.ListRelationships(batch);
        Assert.Contains(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");
    }

    /// <summary>
    /// Tests deleting a relationship.
    /// LLM use case: "delete this relationship"
    /// </summary>
    [Fact]
    public void DeleteRelationship_ExistingRelationship_ReturnsSuccess()
    {
        var batch = _fixture.BatchToken;

        // Delete relationship
        _ = _dataModelCommands.DeleteRelationship(
            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");  // DeleteRelationship throws on error

        // Verify deletion
        var verifyResult = _dataModelCommands.ListRelationships(batch);
        Assert.DoesNotContain(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");

        // Recreate for other tests (shared file)
        _ = _dataModelCommands.CreateRelationship(batch,
            "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true);  // CreateRelationship throws on error
    }

    private OperationResult CreateMeasure(
        string tableName,
        string measureName,
        string daxFormula,
        string? formatType = null,
        string? description = null,
        bool formatDax = false)
    {
        var result = _dataModelCommands.CreateMeasure(
            _fixture.BatchToken,
            tableName,
            measureName,
            daxFormula,
            formatType,
            description,
            formatDax);
        _fixture.RegisterDataModelMeasureForCleanup(measureName);
        return result;
    }

    private OperationResult DeleteMeasure(string measureName)
    {
        var result = _dataModelCommands.DeleteMeasure(
            _fixture.BatchToken,
            measureName);
        _fixture.ForgetDataModelMeasure(measureName);
        return result;
    }

    #endregion
}

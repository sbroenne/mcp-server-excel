using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for Data Model operations focusing on LLM use cases.
/// Tests cover essential workflows: list tables/measures/relationships, create/update/delete measures, manage relationships.
/// Uses DataModelTestsFixture which creates ONE Data Model per test class.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class DataModelCommandsTests : IClassFixture<DataModelTestsFixture>
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly string _dataModelFile;
    private readonly DataModelCreationResult _creationResult;

    /// <summary>
    /// Initializes a new instance of the <see cref="DataModelCommandsTests"/> class.
    /// </summary>
    public DataModelCommandsTests(DataModelTestsFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _dataModelFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
    }

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
        Assert.Equal(3, _creationResult.TablesCreated);
        Assert.Equal(3, _creationResult.TablesLoadedToModel);
        Assert.Equal(2, _creationResult.RelationshipsCreated);
        Assert.Equal(3, _creationResult.MeasuresCreated);
    }

    /// <summary>
    /// Tests listing tables in the data model.
    /// LLM use case: "show me all tables in the data model"
    /// </summary>
    [Fact]
    public async Task ListTables_WithDataModel_ReturnsTables()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.ListTables(batch);

        Assert.True(result.Success, $"ListTables failed: {result.ErrorMessage}");
        Assert.Equal(3, result.Tables.Count);
        Assert.Contains(result.Tables, t => t.Name == "SalesTable");
        Assert.Contains(result.Tables, t => t.Name == "CustomersTable");
        Assert.Contains(result.Tables, t => t.Name == "ProductsTable");
    }

    /// <summary>
    /// Tests getting table details with columns.
    /// LLM use case: "show me the columns in this data model table"
    /// </summary>
    [Fact]
    public async Task GetTable_WithValidTable_ReturnsCompleteInfo()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.ReadTable(batch, "SalesTable");

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
    public async Task GetInfo_WithRealisticDataModel_ReturnsAccurateStatistics()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.ReadInfo(batch);

        Assert.True(result.Success, $"GetModelInfo failed: {result.ErrorMessage}");
        Assert.Equal(3, result.TableCount);
        Assert.Equal(3, result.MeasureCount);
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
    public async Task ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.ListMeasures(batch);

        Assert.True(result.Success, $"ListMeasures failed: {result.ErrorMessage}");
        Assert.NotNull(result.Measures);
        Assert.Equal(3, result.Measures.Count);

        var measureNames = result.Measures.Select(m => m.Name).ToList();
        Assert.Contains("Total Sales", measureNames);
        Assert.Contains("Average Sale", measureNames);
        Assert.Contains("Total Customers", measureNames);
    }

    /// <summary>
    /// Tests viewing a specific measure's DAX formula.
    /// LLM use case: "show me the DAX formula for this measure"
    /// </summary>
    [Fact]
    public async Task Get_WithRealisticDataModel_ReturnsValidDAXFormula()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.Read(batch, "Total Sales");

        Assert.True(result.Success, $"ViewMeasure failed: {result.ErrorMessage}");
        Assert.NotNull(result.DaxFormula);
        Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Amount", result.DaxFormula);
        Assert.Equal("Total Sales", result.MeasureName);
    }

    /// <summary>
    /// Tests creating a new DAX measure.
    /// LLM use case: "create a DAX measure"
    /// </summary>
    [Fact]
    public async Task CreateMeasure_ValidNameAndFormula_CreatesSuccessfully()
    {
        var measureName = $"Test_{nameof(CreateMeasure_ValidNameAndFormula_CreatesSuccessfully)}_{Guid.NewGuid():N}";
        var daxFormula = "SUM(SalesTable[Amount])";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, daxFormula);  // CreateMeasure throws on error

        // Verify measure created
        var listResult = await _dataModelCommands.ListMeasures(batch);
        Assert.Contains(listResult.Measures, m => m.Name == measureName);
    }

    /// <summary>
    /// Tests updating an existing measure's DAX formula.
    /// LLM use case: "update this measure's formula"
    /// </summary>
    [Fact]
    public async Task UpdateMeasure_WithValidFormula_UpdatesSuccessfully()
    {
        var measureName = $"Test_{nameof(UpdateMeasure_WithValidFormula_UpdatesSuccessfully)}_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        var updatedFormula = "AVERAGE(SalesTable[Amount])";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create measure
        _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, originalFormula);  // CreateMeasure throws on error

        // Update formula
        _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: updatedFormula);  // UpdateMeasure throws on error

        // Verify update
        var viewResult = await _dataModelCommands.Read(batch, measureName);
        Assert.Contains("AVERAGE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests deleting a measure.
    /// LLM use case: "delete this DAX measure"
    /// </summary>
    [Fact]
    public async Task DeleteMeasure_WithValidMeasure_ReturnsSuccessResult()
    {
        var measureName = $"Test_{nameof(DeleteMeasure_WithValidMeasure_ReturnsSuccessResult)}_{Guid.NewGuid():N}";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create measure
        _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, "SUM(SalesTable[Amount])");  // CreateMeasure throws on error

        // Delete measure
        _dataModelCommands.DeleteMeasure(batch, measureName);  // DeleteMeasure throws on error

        // Verify deletion
        var listResult = await _dataModelCommands.ListMeasures(batch);
        Assert.DoesNotContain(listResult.Measures, m => m.Name == measureName);
    }

    #endregion

    #region Relationship Operations (3 tests)

    /// <summary>
    /// Tests listing all relationships in the data model.
    /// LLM use case: "show me all table relationships"
    /// </summary>
    [Fact]
    public async Task ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithDetails()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.ListRelationships(batch);

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
    public async Task CreateRelationship_ValidTablesAndColumns_CreatesSuccessfully()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Delete existing relationship first to allow recreating it
        var listResult = await _dataModelCommands.ListRelationships(batch);
        if (listResult.Success && listResult.Relationships?.Any(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID") == true)
        {
            _dataModelCommands.DeleteRelationship(batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");  // DeleteRelationship throws on error
        }

        // Create relationship
        _dataModelCommands.CreateRelationship(
            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");  // CreateRelationship throws on error

        // Verify creation
        var verifyResult = await _dataModelCommands.ListRelationships(batch);
        Assert.Contains(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");
    }

    /// <summary>
    /// Tests deleting a relationship.
    /// LLM use case: "delete this relationship"
    /// </summary>
    [Fact]
    public async Task DeleteRelationship_ExistingRelationship_ReturnsSuccess()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Delete relationship
        _dataModelCommands.DeleteRelationship(
            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");  // DeleteRelationship throws on error

        // Verify deletion
        var verifyResult = await _dataModelCommands.ListRelationships(batch);
        Assert.DoesNotContain(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");

        // Recreate for other tests (shared file)
        _dataModelCommands.CreateRelationship(batch,
            "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true);  // CreateRelationship throws on error
    }

    #endregion
}

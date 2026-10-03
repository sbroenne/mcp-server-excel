using System.Globalization;
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
/// Uses a saved Data Model template with an isolated workbook for each test.
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
        Assert.True(File.Exists(_dataModelFile));
        Assert.Equal(
            ["CustomersTable", "DisambiguationTable", "ProductsTable", "RegionalSalesTable", "SalesTable"],
            RequireSuccess(_dataModelCommands.ListTables(_fixture.BatchToken))
                .Tables.Select(table => table.Name).Order(StringComparer.Ordinal));
        Assert.Equal(2455d, ReadMeasureValue("Total Sales"));
        Assert.Equal(5d, ReadMeasureValue("Total Customers"));
    }

    /// <summary>
    /// Tests listing tables in the data model.
    /// LLM use case: "show me all tables in the data model"
    /// </summary>
    [Fact]
    public void ListTables_WithDataModel_ReturnsTables()
    {
        var batch = _fixture.BatchToken;
        var result = RequireSuccess(_dataModelCommands.ListTables(batch));

        Assert.Equal(5, result.Tables.Count);  // Now includes RegionalSalesTable and DisambiguationTable
        Assert.Equal(
            ["CustomersTable", "DisambiguationTable", "ProductsTable", "RegionalSalesTable", "SalesTable"],
            result.Tables.Select(table => table.Name).Order(StringComparer.Ordinal));
    }

    /// <summary>
    /// Tests getting table details with columns.
    /// LLM use case: "show me the columns in this data model table"
    /// </summary>
    [Fact]
    public void GetTable_WithValidTable_ReturnsCompleteInfo()
    {
        var batch = _fixture.BatchToken;
        var result = RequireSuccess(_dataModelCommands.ReadTable(batch, "SalesTable"));

        Assert.Equal("SalesTable", result.TableName);
        Assert.Equal(10, result.RecordCount);
        Assert.Equal(
            ["Amount", "CustomerID", "Date", "ProductID", "Quantity", "SalesID"],
            result.Columns.Select(column => column.Name).Order(StringComparer.Ordinal));
        Assert.Equal("WORKSHEET", result.SourceConnectionType);
        Assert.Equal("Excel Table: SalesTable", result.SourceConnectionDescription);
        Assert.Equal(2455d, ReadMeasureValue("Total Sales"));
    }

    /// <summary>
    /// Tests getting data model statistics.
    /// LLM use case: "show me information about this data model"
    /// </summary>
    [Fact]
    public void GetInfo_WithRealisticDataModel_ReturnsAccurateStatistics()
    {
        var batch = _fixture.BatchToken;
        var result = RequireSuccess(_dataModelCommands.ReadInfo(batch));

        Assert.Equal(5, result.TableCount);  // Now includes RegionalSalesTable and DisambiguationTable
        Assert.Equal(6, result.MeasureCount);  // Now includes TotalRevenue, ACR, Discount
        Assert.Equal(2, result.RelationshipCount);
        Assert.Equal(31, result.TotalRows);
        Assert.Equal(
            ["CustomersTable", "DisambiguationTable", "ProductsTable", "RegionalSalesTable", "SalesTable"],
            result.TableNames.Order(StringComparer.Ordinal));
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
        var result = RequireSuccess(_dataModelCommands.ListMeasures(batch));

        Assert.Equal(6, result.Measures.Count);  // Now includes TotalRevenue, ACR, Discount

        Assert.Equal(
            ["ACR", "Average Sale", "Discount", "Total Customers", "Total Sales", "TotalRevenue"],
            result.Measures.Select(measure => measure.Name).Order(StringComparer.Ordinal));
        Assert.Equal("SUM(SalesTable[Amount])",
            Assert.Single(result.Measures, measure => measure.Name == "Total Sales").FormulaPreview);
        Assert.Equal(2455d, ReadMeasureValue("Total Sales"));
    }

    /// <summary>
    /// Tests viewing a specific measure's DAX formula.
    /// LLM use case: "show me the DAX formula for this measure"
    /// </summary>
    [Fact]
    public void Get_WithRealisticDataModel_ReturnsValidDAXFormula()
    {
        var batch = _fixture.BatchToken;
        var result = RequireSuccess(_dataModelCommands.Read(batch, "Total Sales"));

        Assert.Equal("SUM(SalesTable[Amount])", result.DaxFormula);
        Assert.Equal("Total Sales", result.MeasureName);
        Assert.Equal(2455d, ReadMeasureValue("Total Sales"));
    }

    /// <summary>
    /// Tests that Read returns structured FormatInfo with type and properties.
    /// Regression test: FormatString was previously a string, now it's a structured FormatInfo object.
    /// </summary>
    [Fact]
    public void Read_WithMeasure_ReturnsStructuredFormatInfo()
    {
        var batch = _fixture.BatchToken;
        var result = RequireSuccess(_dataModelCommands.Read(batch, "Total Sales"));

        var format = Assert.IsType<MeasureFormatInfo>(result.FormatInfo);
        Assert.Equal("Decimal", format.Type);
        Assert.Equal(0, format.DecimalPlaces);
        Assert.Equal(2455d, ReadMeasureValue("Total Sales"));
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
        RequireSuccess(CreateMeasure("SalesTable", measureName, daxFormula));

        var listResult = RequireSuccess(_dataModelCommands.ListMeasures(batch));
        Assert.Contains(listResult.Measures, m => m.Name == measureName);
        Assert.Equal(daxFormula, RequireSuccess(_dataModelCommands.Read(batch, measureName)).DaxFormula);
        Assert.Equal(2455d, ReadMeasureValue(measureName));
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
        RequireSuccess(CreateMeasure("SalesTable", measureName, originalFormula));

        RequireSuccess(_dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: updatedFormula));

        var viewResult = RequireSuccess(_dataModelCommands.Read(batch, measureName));
        Assert.Equal(updatedFormula, viewResult.DaxFormula);
        Assert.Equal(245.5d, ReadMeasureValue(measureName));
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
        RequireSuccess(CreateMeasure("SalesTable", measureName, "SUM(SalesTable[Amount])"));

        RequireSuccess(DeleteMeasure(measureName));

        var listResult = RequireSuccess(_dataModelCommands.ListMeasures(batch));
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
        var result = RequireSuccess(_dataModelCommands.ListRelationships(batch));

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
        var listResult = RequireSuccess(_dataModelCommands.ListRelationships(batch));
        if (listResult.Relationships.Any(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID"))
        {
            RequireSuccess(_dataModelCommands.DeleteRelationship(
                batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID"));
        }

        RequireSuccess(_dataModelCommands.CreateRelationship(
            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID"));

        var verifyResult = RequireSuccess(_dataModelCommands.ListRelationships(batch));
        Assert.Contains(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID" && r.IsActive);
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
        var before = RequireSuccess(_dataModelCommands.ListRelationships(batch));
        Assert.Equal(2, before.Relationships.Count);
        RequireSuccess(_dataModelCommands.DeleteRelationship(
            batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID"));

        var verifyResult = RequireSuccess(_dataModelCommands.ListRelationships(batch));
        Assert.Single(verifyResult.Relationships);
        Assert.DoesNotContain(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");

        RequireSuccess(_dataModelCommands.CreateRelationship(batch,
            "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true));
        var recovered = RequireSuccess(_dataModelCommands.ListRelationships(batch));
        Assert.Equal(2, recovered.Relationships.Count);
        Assert.Contains(recovered.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID" && r.IsActive);
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

    private double ReadMeasureValue(string measureName)
    {
        var result = RequireSuccess(_dataModelCommands.Evaluate(
            _fixture.BatchToken, $"EVALUATE ROW(\"Value\", [{measureName}])"));
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal("[Value]", Assert.Single(result.Columns));
        return Convert.ToDouble(Assert.Single(Assert.Single(result.Rows)), CultureInfo.InvariantCulture);
    }

    #endregion
}

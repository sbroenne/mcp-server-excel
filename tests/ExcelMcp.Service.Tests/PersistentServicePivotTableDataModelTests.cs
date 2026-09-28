using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for PivotTable creation from Power Pivot Data Model tables.
/// Uses DataModelPivotTableFixture which creates ONE comprehensive Data Model + PivotTable workbook (shared via collection fixture).
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Feature", "PivotTables")]
[Trait("Speed", "Slow")]
public class PersistentServicePivotTableDataModelTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IPivotTableCommands _pivotCommands =
        fixture.CreateCommands<IPivotTableCommands>();

    /// <summary>
    /// Tests creating PivotTable from Data Model table.
    /// </summary>
    [Fact]
    public void CreateFromDataModel_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange - Use shared Data Model fixture
        Assert.True(
            PersistentServiceDataModelFixture.CreationResult.Success,
            "Data Model fixture must be created successfully");

        // Act - Create PivotTable from Data Model table
        var result = _pivotCommands.CreateFromDataModel(
            _fixture.BatchToken,
            "SalesTable",  // Data Model table name from fixture
            "SalesData",   // Destination sheet (exists in fixture)
            "H1",          // Destination cell
            "SalesDataModelPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesDataModelPivot", result.PivotTableName);
        Assert.Equal("SalesData", result.SheetName);
        Assert.NotEmpty(result.Range);
        Assert.Contains("ThisWorkbookDataModel", result.SourceData);
        Assert.True(result.SourceRowCount > 0, "Should have rows in source Data Model table");
        Assert.NotEmpty(result.AvailableFields);

        // Verify expected fields from SalesTable in Data Model
        Assert.Contains("SalesID", result.AvailableFields);
        Assert.Contains("CustomerID", result.AvailableFields);
        Assert.Contains("Amount", result.AvailableFields);
    }

    /// <summary>
    /// Tests error handling for non-existent Data Model table.
    /// </summary>
    [Fact]
    public void CreateFromDataModel_NonExistentTable_ReturnsError()
    {
        // Arrange
        Assert.True(
            PersistentServiceDataModelFixture.CreationResult.Success,
            "Data Model fixture must be created successfully");

        // Act & Assert - Try to create PivotTable from non-existent table (should throw)
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _pivotCommands.CreateFromDataModel(
                _fixture.BatchToken,
                "NonExistentTable",
                "SalesData",
                "H1",
                "FailedPivot"));

        Assert.Contains("not found in Data Model", exception.Message);
    }

    /// <summary>
    /// Tests that all fields from Data Model table are discovered.
    /// </summary>
    [Fact]
    public void CreateFromDataModel_MultipleFieldsAvailable_ReturnsAllColumns()
    {
        // Arrange
        Assert.True(
            PersistentServiceDataModelFixture.CreationResult.Success,
            "Data Model fixture must be created successfully");

        // Act - Create PivotTable and verify all fields are discovered
        var result = _pivotCommands.CreateFromDataModel(
            _fixture.BatchToken,
            "CustomersTable",  // Has 4 columns: CustomerID, Name, Region, Country
            "Customers",
            "H1",
            "CustomersPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal(4, result.AvailableFields.Count);
        Assert.Contains("CustomerID", result.AvailableFields);
        Assert.Contains("Name", result.AvailableFields);
        Assert.Contains("Region", result.AvailableFields);
        Assert.Contains("Country", result.AvailableFields);
    }
}




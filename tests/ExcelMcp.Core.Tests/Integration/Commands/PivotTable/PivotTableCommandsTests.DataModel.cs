using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Integration tests for PivotTable creation from Power Pivot Data Model tables.
/// Uses DataModelTestsFixture which creates ONE Data Model file per test class.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Feature", "PivotTables")]
[Trait("Speed", "Slow")]
public class PivotTableDataModelTests : IClassFixture<DataModelTestsFixture>
{
    private readonly PivotTableCommands _pivotCommands;
    private readonly string _dataModelFile;
    private readonly DataModelCreationResult _creationResult;

    /// <summary>
    /// Initializes a new instance of the <see cref="PivotTableDataModelTests"/> class.
    /// </summary>
    public PivotTableDataModelTests(DataModelTestsFixture fixture)
    {
        _pivotCommands = new PivotTableCommands();
        _dataModelFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
    }

    /// <summary>
    /// Tests creating PivotTable from Data Model table.
    /// </summary>
    [Fact]
    public async Task CreateFromDataModel_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange - Use shared Data Model fixture
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        // Act - Create PivotTable from Data Model table
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _pivotCommands.CreateFromDataModel(
            batch,
            "SalesTable",  // Data Model table name from fixture
            "Sales",       // Destination sheet
            "H1",          // Destination cell
            "DataModelPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("DataModelPivot", result.PivotTableName);
        Assert.Equal("Sales", result.SheetName);
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
    public async Task CreateFromDataModel_NonExistentTable_ReturnsError()
    {
        // Arrange
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        // Act - Try to create PivotTable from non-existent table
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _pivotCommands.CreateFromDataModel(
            batch,
            "NonExistentTable",
            "Sales",
            "H1",
            "FailedPivot");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found in Data Model", result.ErrorMessage);
    }

    /// <summary>
    /// Tests that all fields from Data Model table are discovered.
    /// </summary>
    [Fact]
    public async Task CreateFromDataModel_MultipleFieldsAvailable_ReturnsAllColumns()
    {
        // Arrange
        Assert.True(_creationResult.Success, "Data Model fixture must be created successfully");

        // Act - Create PivotTable and verify all fields are discovered
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _pivotCommands.CreateFromDataModel(
            batch,
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

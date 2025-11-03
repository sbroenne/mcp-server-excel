using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable creation operations
/// </summary>
public partial class PivotTableCommandsTests
{
    [Fact]
    public async Task CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.Equal("SalesData", result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
    }

    [Fact]
    public async Task CreateFromTable_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(CreateFromTable_WithValidTable_CreatesCorrectPivotStructure));

        // Act - Use single batch for table creation and pivot creation
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create table first
        var tableCommands = new TableCommands();
        var tableResult = await tableCommands.CreateAsync(batch, "SalesData", "SalesTable", "A1:D6", true, "TableStyleMedium2");
        Assert.True(tableResult.Success, $"Table creation failed: {tableResult.ErrorMessage}");

        // Create pivot from table
        var result = await _pivotCommands.CreateFromTableAsync(
            batch,
            "SalesTable",
            "SalesData", "F1",
            "TablePivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TablePivot", result.PivotTableName);
        Assert.Equal("SalesData", result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
    }

    [Fact]
    [Trait("Feature", "DataModel")]
    public async Task CreateFromDataModel_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange - Use DataModelTestsFixture file that has Data Model tables
        var dataModelFixture = new Helpers.DataModelTestsFixture();
        await dataModelFixture.InitializeAsync();
        
        try
        {
            var testFile = dataModelFixture.TestFilePath;

            // Act - Create PivotTable from Data Model table
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var result = await _pivotCommands.CreateFromDataModelAsync(
                batch,
                "SalesTable",  // Data Model table name
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
        finally
        {
            await dataModelFixture.DisposeAsync();
        }
    }

    [Fact]
    [Trait("Feature", "DataModel")]
    public async Task CreateFromDataModel_NonExistentTable_ReturnsError()
    {
        // Arrange
        var dataModelFixture = new Helpers.DataModelTestsFixture();
        await dataModelFixture.InitializeAsync();
        
        try
        {
            var testFile = dataModelFixture.TestFilePath;

            // Act - Try to create PivotTable from non-existent table
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var result = await _pivotCommands.CreateFromDataModelAsync(
                batch,
                "NonExistentTable",
                "Sales",
                "H1",
                "FailedPivot");

            // Assert
            Assert.False(result.Success);
            Assert.Contains("not found in Data Model", result.ErrorMessage);
        }
        finally
        {
            await dataModelFixture.DisposeAsync();
        }
    }

    [Fact]
    [Trait("Feature", "DataModel")]
    public async Task CreateFromDataModel_NoDataModel_ReturnsError()
    {
        // Arrange - Use regular file without Data Model
        var testFile = await CreateTestFileWithDataAsync(nameof(CreateFromDataModel_NoDataModel_ReturnsError));

        // Act - Try to create PivotTable from Data Model when none exists
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.CreateFromDataModelAsync(
            batch,
            "AnyTable",
            "SalesData",
            "F1",
            "FailedPivot");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("does not contain a Power Pivot Data Model", result.ErrorMessage);
    }

    [Fact]
    public async Task AddRowField_WithValidField_AddsFieldToRows()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(AddRowField_WithValidField_AddsFieldToRows));

        // Act - Use single batch for create and add field
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create pivot
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add row field
        var result = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

    [Fact]
    public async Task ListFields_AfterCreate_ReturnsAvailableFields()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(ListFields_AfterCreate_ReturnsAvailableFields));

        // Act - Use single batch for create and list fields
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create pivot
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // List fields
        var result = await _pivotCommands.ListFieldsAsync(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Fields);
        Assert.True(result.Fields.Count >= 4); // Region, Product, Sales, Date
    }
}

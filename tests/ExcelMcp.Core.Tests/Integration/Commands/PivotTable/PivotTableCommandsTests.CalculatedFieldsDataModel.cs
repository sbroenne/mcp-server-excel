using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for calculated fields with Data Model / OLAP PivotTables.
/// OLAP PivotTables do NOT support CalculatedFields (Excel COM limitation).
/// For OLAP, use DAX measures via excel_datamodel tool instead.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public class PivotTableCalculatedFieldsDataModelTests : IClassFixture<DataModelTestsFixture>
{
    private readonly PivotTableCommands _pivotCommands;
    private readonly string _dataModelFile;
    private readonly DataModelCreationResult _creationResult;

    public PivotTableCalculatedFieldsDataModelTests(DataModelTestsFixture fixture)
    {
        _pivotCommands = new PivotTableCommands();
        _dataModelFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
    }

    [Fact]
    public void CreateCalculatedField_OlapPivotTable_ReturnsNotSupported()
    {
        // Arrange - Verify Data Model exists
        Assert.True(_creationResult.Success, $"Data Model creation failed: {_creationResult.ErrorMessage}");

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create OLAP PivotTable from Data Model (SalesTable has: SalesID, Date, CustomerID, ProductID, Amount, Quantity)
        // Use existing "Sales" sheet from fixture
        var createResult = _pivotCommands.CreateFromDataModel(
            batch, "SalesTable", "Sales", "K1", "OlapSalesCalcTest");
        Assert.True(createResult.Success, $"Failed to create OLAP PivotTable: {createResult.ErrorMessage}");

        // Add fields to PivotTable
        var rowResult = _pivotCommands.AddRowField(batch, "OlapSalesCalcTest", "ProductID");
        Assert.True(rowResult.Success, $"AddRowField failed: {rowResult.ErrorMessage}");

        var valueResult = _pivotCommands.AddValueField(batch, "OlapSalesCalcTest", "Amount");
        Assert.True(valueResult.Success, $"AddValueField failed: {valueResult.ErrorMessage}");

        // Act - Attempt to create calculated field on OLAP PivotTable
        var result = _pivotCommands.CreateCalculatedField(batch, "OlapSalesCalcTest", "TestField", "=Amount*2");

        // Assert - Should fail with NotSupported message
        Assert.False(result.Success, "CreateCalculatedField should fail for OLAP PivotTables");
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not supported", result.ErrorMessage.ToLowerInvariant());
        Assert.Contains("OLAP", result.ErrorMessage);

        // Verify workflow hint points to DAX measures
        Assert.NotNull(result.WorkflowHint);
        Assert.Contains("excel_datamodel", result.WorkflowHint);
        Assert.Contains("DAX", result.WorkflowHint);
    }
}

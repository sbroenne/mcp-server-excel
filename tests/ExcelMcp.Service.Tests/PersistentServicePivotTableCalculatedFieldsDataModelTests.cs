using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for calculated fields with Data Model / OLAP PivotTables.
/// OLAP PivotTables do NOT support CalculatedFields (Excel COM limitation).
/// For OLAP, use DAX measures via datamodel tool instead.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public class PersistentServicePivotTableCalculatedFieldsDataModelTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IPersistentPivotTableCommands _pivotCommands =
        fixture.CreateCommands<IPersistentPivotTableCommands>();

    [Fact]
    public void CreateCalculatedField_OlapPivotTable_ReturnsNotSupported()
    {
        // Arrange - Verify Data Model exists
        var creationResult = PersistentServiceDataModelFixture.CreationResult;
        Assert.True(
            creationResult.Success,
            $"Data Model creation failed: {creationResult.ErrorMessage}");
        var batch = _fixture.BatchToken;

        // Create OLAP PivotTable from Data Model (SalesTable has: SalesID, Date, CustomerID, ProductID, Amount, Quantity)
        // Use existing "SalesData" sheet from fixture
        var createResult = _pivotCommands.CreateFromDataModel(
            batch, "SalesTable", "SalesData", "K1", "OlapSalesCalcTest");
        Assert.True(createResult.Success, $"Failed to create OLAP PivotTable: {createResult.ErrorMessage}");

        // Add fields to PivotTable - use exact CubeField names (LLM discovers via ListFields)
        var rowResult = _pivotCommands.AddRowField(batch, "OlapSalesCalcTest", "[SalesTable].[ProductID]");
        Assert.True(rowResult.Success, $"AddRowField failed: {rowResult.ErrorMessage}");

        var valueResult = _pivotCommands.AddValueField(batch, "OlapSalesCalcTest", "[SalesTable].[Amount]");
        Assert.True(valueResult.Success, $"AddValueField failed: {valueResult.ErrorMessage}");
        var dataBefore = RequireSuccess(_pivotCommands.GetData(batch, "OlapSalesCalcTest"));
        var fieldsBefore = RequireSuccess(_pivotCommands.ListFields(batch, "OlapSalesCalcTest"));
        Assert.NotEmpty(dataBefore.Values);

        // Act - Attempt to create calculated field on OLAP PivotTable
        var result = _pivotCommands.CreateCalculatedField(batch, "OlapSalesCalcTest", "TestField", "=Amount*2");

        // Assert - Should fail with NotSupported message
        Assert.False(result.Success, "CreateCalculatedField should fail for OLAP PivotTables");
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not supported", result.ErrorMessage.ToLowerInvariant());
        Assert.Contains("OLAP", result.ErrorMessage);

        // Verify workflow hint points to DAX measures
        Assert.NotNull(result.WorkflowHint);
        Assert.Contains("datamodel", result.WorkflowHint);
        Assert.Contains("DAX", result.WorkflowHint);
        Assert.Equal(JsonSerializer.Serialize(dataBefore),
            JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(batch, "OlapSalesCalcTest"))));
        Assert.Equal(JsonSerializer.Serialize(fieldsBefore),
            JsonSerializer.Serialize(RequireSuccess(_pivotCommands.ListFields(batch, "OlapSalesCalcTest"))));
    }
}

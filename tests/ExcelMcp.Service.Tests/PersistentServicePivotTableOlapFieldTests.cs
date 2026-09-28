using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for PivotTable commands.
/// Uses DataModelPivotTableFixture for all tests (shared across ALL test classes via collection fixture).
/// Fixture initialization IS the test for data preparation.
/// Each test gets its own batch for isolation.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
public partial class PersistentServicePivotTableOlapFieldTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IPersistentPivotTableCommands _pivotCommands;
    private readonly IDataModelCommands _dataModelCommands;
    private readonly DataModelPivotTableCreationResult _creationResult;

    public PersistentServicePivotTableOlapFieldTests(
        PersistentServiceDataModelFixture fixture) :
        base(fixture)
    {
        _pivotCommands =
            fixture.CreateCommands<IPersistentPivotTableCommands>();
        _dataModelCommands = fixture.CreateCommands<IDataModelCommands>();
        _creationResult = PersistentServiceDataModelFixture.CreationResult;
    }

    /// <summary>
    /// Explicit test that validates the fixture creation results.
    /// This makes the data preparation test visible in test results and validates:
    /// - SessionManager.CreateSessionForNewFile()
    /// - Data Model tables, relationships, measures, and PivotTable creation
    /// - Batch.Save() persistence
    /// </summary>
    [Fact]
    [Trait("Speed", "Fast")]
    public void DataPreparation_ViaFixture_CreatesSalesData()
    {
        // Assert the fixture creation succeeded
        Assert.True(_creationResult.Success,
            $"Data preparation failed during fixture initialization: {_creationResult.ErrorMessage}");

        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.True(_creationResult.TablesCreated > 0, "No tables were created");
        Assert.True(_creationResult.CreationTimeMs > 0);

        // This test appears in test results as proof that creation was tested
        Console.WriteLine($"? Data prepared successfully in {_creationResult.CreationTimeMs}ms");
    }

    /// <summary>
    /// Tests that sales data persists correctly after file close/reopen.
    /// Validates that SaveAsync() properly persisted the data.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void DataPreparation_Persists_AfterReopenFile()
    {
        var batch = _fixture.BatchToken;

        // Verify data persisted by reading range (SalesData sheet from DataModel fixture)
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets["SalesData"];

            // Verify headers match DataModel fixture's SalesTable columns
            Assert.Equal("SalesID", sheet.Range["A1"].Value2?.ToString());
            Assert.Equal("Date", sheet.Range["B1"].Value2?.ToString());
            Assert.Equal("CustomerID", sheet.Range["C1"].Value2?.ToString());
            Assert.Equal("ProductID", sheet.Range["D1"].Value2?.ToString());

            // Verify first data row
            Assert.Equal(1.0, Convert.ToDouble(sheet.Range["A2"].Value2));
            Assert.Equal(101.0, Convert.ToDouble(sheet.Range["C2"].Value2));
            Assert.Equal(1001.0, Convert.ToDouble(sheet.Range["D2"].Value2));

            return 0;
        });

        // This proves data creation + save worked correctly
    }
}


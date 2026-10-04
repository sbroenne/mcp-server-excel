using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for PivotTable commands.
/// Uses a saved Data Model template and a shared workbook session.
/// Each field test creates an isolated PivotTable; preparation checks verify the saved template.
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
        Assert.Equal(5, _creationResult.TablesCreated);
        Assert.True(_creationResult.CreationTimeMs > 0);

        var data = _commands.GetValues(_fixture.BatchToken, "RegionalData", "A1:D9");
        Assert.True(data.Success, data.ErrorMessage);
        Assert.Equal(9, data.RowCount);
        Assert.Equal(4, data.ColumnCount);
        Assert.Equal(["Quarter", "Region", "Sales", "Units"],
            data.Values[0].Select(value => value?.ToString()));
        (string Quarter, string Region, double Sales, double Units)[] expected =
        [
            ("Q1", "North", 5000, 100), ("Q1", "South", 6000, 120),
            ("Q1", "East", 5500, 110), ("Q1", "West", 7000, 140),
            ("Q2", "North", 5500, 110), ("Q2", "South", 6500, 130),
            ("Q2", "East", 6000, 120), ("Q2", "West", 7500, 150)
        ];
        Assert.Equal(expected.Length + 1, data.Values.Count);
        for (var index = 0; index < expected.Length; index++)
        {
            var row = data.Values[index + 1];
            Assert.Equal(4, row.Count);
            Assert.Equal(expected[index].Quarter, row[0]);
            Assert.Equal(expected[index].Region, row[1]);
            Assert.Equal(expected[index].Sales, Convert.ToDouble(row[2], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(expected[index].Units, Convert.ToDouble(row[3], System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    /// <summary>
    /// Tests that sales data persists correctly after file close/reopen.
    /// Validates that SaveAsync() properly persisted the data.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task DataPreparation_Persists_AfterReopenFile()
    {
        await _fixture.SaveAndReopenAsync();
        var batch = _fixture.BatchToken;
        var data = _commands.GetValues(batch, "SalesData", "A1:F2");
        Assert.True(data.Success, data.ErrorMessage);
        Assert.Equal(["SalesID", "Date", "CustomerID", "ProductID", "Amount", "Quantity"],
            data.Values[0].Select(value => value?.ToString()));
        Assert.Equal(1, Convert.ToInt32(data.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(new DateTime(2024, 1, 15), DateTime.FromOADate(
            Convert.ToDouble(data.Values[1][1], System.Globalization.CultureInfo.InvariantCulture)));
        Assert.Equal(101, Convert.ToInt32(data.Values[1][2], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(1001, Convert.ToInt32(data.Values[1][3], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(150d, Convert.ToDouble(data.Values[1][4], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(2, Convert.ToInt32(data.Values[1][5], System.Globalization.CultureInfo.InvariantCulture));
    }
}

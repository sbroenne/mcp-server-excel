using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for PivotTable creation from Power Pivot Data Model tables.
/// Uses a class-scoped Data Model fixture and test-owned destination sheets.
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
    private static readonly string[] SalesFields = ["SalesID", "Date", "CustomerID", "ProductID", "Amount", "Quantity"];
    private static readonly string[] CustomerFields = ["CustomerID", "Name", "Region", "Country"];
    private static readonly string[] CustomerNames = ["Acme Corp", "Beta Inc", "Delta Co", "Epsilon Ltd", "Gamma LLC"];
    private readonly IPersistentPivotTableCommands _pivotCommands =
        fixture.CreateCommands<IPersistentPivotTableCommands>();

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

        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var result = _pivotCommands.CreateFromDataModel(
            _fixture.BatchToken,
            "SalesTable",  // Data Model table name from fixture
            sheet,
            "H1",          // Destination cell
            "SalesDataModelPivot");

        // Assert
        RequireSuccess(result);
        Assert.Equal("SalesDataModelPivot", result.PivotTableName);
        Assert.Equal(sheet, result.SheetName);
        Assert.NotEmpty(result.Range);
        Assert.Equal("ThisWorkbookDataModel[SalesTable]", result.SourceData);
        Assert.Equal(10, result.SourceRowCount);
        Assert.Equal(SalesFields,
            result.AvailableFields);
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, result.PivotTableName,
            "[Measures].[Total Sales]", AggregationFunction.Sum, "Sales"));
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, result.PivotTableName, null));
        AssertNativeData(sheet, result.PivotTableName, 2455);
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

        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "H1", [["retained"]]));
        var pivotsBefore = RequireSuccess(_pivotCommands.List(_fixture.BatchToken));
        var model = _fixture.Send("datamodel.evaluate", new { daxQuery = "EVALUATE SalesTable" });
        Assert.True(model.Success, model.ErrorMessage);
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _pivotCommands.CreateFromDataModel(
                _fixture.BatchToken,
                "NonExistentTable",
                sheet,
                "H1",
                "FailedPivot"));

        Assert.Contains("not found in Data Model", exception.Message);
        Assert.Equal("retained", RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "H1")).Values[0][0]);
        var pivotsAfter = RequireSuccess(_pivotCommands.List(_fixture.BatchToken));
        Assert.Equal(pivotsBefore.PivotTables.Select(p => p.Name), pivotsAfter.PivotTables.Select(p => p.Name));
        var unchanged = _fixture.Send("datamodel.evaluate", new { daxQuery = "EVALUATE SalesTable" });
        Assert.True(unchanged.Success, unchanged.ErrorMessage);
        Assert.Equal(model.Result, unchanged.Result);
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

        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var result = _pivotCommands.CreateFromDataModel(
            _fixture.BatchToken,
            "CustomersTable",  // Has 4 columns: CustomerID, Name, Region, Country
            sheet,
            "H1",
            "CustomersPivot");

        // Assert
        RequireSuccess(result);
        Assert.Equal(5, result.SourceRowCount);
        Assert.Equal("ThisWorkbookDataModel[CustomersTable]", result.SourceData);
        Assert.Equal(CustomerFields, result.AvailableFields);
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, result.PivotTableName,
            "[CustomersTable].[Name]", null));
        var data = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, result.PivotTableName));
        Assert.Equal(CustomerNames,
            data.Values.Skip(1).Take(5).Select(row => row[0]));
    }

    private void AssertNativeData(string sheetName, string pivotName, double expected)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.Range? data = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cache = pivot.PivotCache();
                Assert.True(cache.OLAP);
                data = pivot.DataBodyRange;
                Assert.Equal(expected, Convert.ToDouble(data.Value2,
                    System.Globalization.CultureInfo.InvariantCulture), 8);
            }
            finally
            {
                ComUtilities.Release(ref data);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}


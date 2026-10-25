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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShareExistingCache_DataModelOnEitherSide_IsRejectedBeforeWrites(bool modelIsMoved)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B3",
            [["Region", "Sales"], ["North", 100], ["South", 200]]));
        RequireSuccess(_pivotCommands.CreateFromRange(batch, sheet, "A1:B3", sheet, "D1", "WorksheetPivot"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "WorksheetPivot", "Sales"));
        RequireSuccess(_pivotCommands.CreateFromDataModel(batch, "SalesTable", sheet, "H1", "ModelPivot"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "ModelPivot", "[Measures].[Total Sales]"));
        RequireSuccess(_pivotCommands.Refresh(batch, "ModelPivot"));
        var before = RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Select(pivot => (pivot.SheetName, pivot.Name, pivot.CacheIndex)).ToArray();

        var error = Assert.Throws<NotSupportedException>(() => _pivotCommands.ShareCache(batch,
            sheet, modelIsMoved ? "ModelPivot" : "WorksheetPivot",
            modelIsMoved ? "WorksheetPivot" : "ModelPivot", sheet));

        Assert.Contains("worksheet-backed", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Select(pivot => (pivot.SheetName, pivot.Name, pivot.CacheIndex)).ToArray());
        AssertNativeData(sheet, "ModelPivot", 2455);
        var values = RequireSuccess(_pivotCommands.GetData(batch, "WorksheetPivot")).Values;
        Assert.Equal(300d, Convert.ToDouble(values[^1][0], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedCache_DataModelTarget_IsRejectedBeforeCreatingReportOrCache(bool table)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B3",
            [["Region", "Sales"], ["North", 100], ["South", 200]]));
        var tables = fixture.CreateCommands<IPersistentTableCommands>();
        RequireSuccess(tables.Create(batch, sheet, "WorksheetSales", "A1:B3", true));
        RequireSuccess(_pivotCommands.CreateFromDataModel(batch, "SalesTable", sheet, "H1", "ModelPivot"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "ModelPivot", "[Measures].[Total Sales]"));
        RequireSuccess(_pivotCommands.Refresh(batch, "ModelPivot"));
        AssertNativeData(sheet, "ModelPivot", 2455);
        var before = ReadCacheState();
        var reportsBefore = RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Select(pivot => (pivot.SheetName, pivot.Name, pivot.CacheIndex)).ToArray();

        var error = Assert.Throws<NotSupportedException>(() =>
        {
            if (table)
                _pivotCommands.CreateFromTable(batch, "WorksheetSales", sheet, "P1", "Rejected", "ModelPivot", sheet);
            else
                _pivotCommands.CreateFromRange(batch, sheet, "A1:B3", sheet, "P1", "Rejected", "ModelPivot", sheet);
        });

        Assert.Contains("worksheet-backed", error.Message, StringComparison.Ordinal);
        Assert.Contains("OLAP/Data Model", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, ReadCacheState());
        var reportsAfter = RequireSuccess(_pivotCommands.List(batch)).PivotTables;
        Assert.Equal(reportsBefore, reportsAfter.Select(pivot => (pivot.SheetName, pivot.Name, pivot.CacheIndex)).ToArray());
        var listed = Assert.Single(reportsAfter, pivot => pivot.SheetName == sheet);
        Assert.Equal("ModelPivot", listed.Name);
        Assert.Equal(before.Index, listed.CacheIndex);
        Assert.Equal(before.Index, RequireSuccess(_pivotCommands.Read(batch, "ModelPivot", sheet)).PivotTable.CacheIndex);
        Assert.True(before.Index > 0);
        AssertNativeData(sheet, "ModelPivot", 2455);
        Assert.Equal(new object?[] { "Region", "Sales" },
            RequireSuccess(_commands.GetValues(batch, sheet, "A1:B3")).Values[0]);

        (int Count, int Index) ReadCacheState() => _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.PivotCaches? caches = null;
            Excel.Worksheet? nativeSheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            try
            {
                caches = context.Book.PivotCaches();
                nativeSheet = ComUtilities.FindSheet(context.Book, sheet);
                pivots = nativeSheet!.PivotTables();
                pivot = pivots.Item("ModelPivot");
                return (caches.Count, pivot.CacheIndex);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref nativeSheet);
                ComUtilities.Release(ref caches);
            }
        });
    }

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

    [Fact]
    public void CreateFromDataModel_DuplicateNameReportsExistingPivotAndPreservesIt()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"DuplicatePivot{Guid.NewGuid():N}";
        RequireSuccess(_pivotCommands.CreateFromDataModel(batch, "SalesTable", sheet, "H1", pivotName));
        var before = RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Select(pivot => (pivot.SheetName, pivot.Name, pivot.CacheIndex))
            .ToArray();
        var original = Assert.Single(before, pivot => pivot.Name == pivotName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _pivotCommands.CreateFromDataModel(batch, "SalesTable", sheet, "P1", pivotName));

        Assert.Contains($"PivotTable '{pivotName}' already exists", exception.Message, StringComparison.Ordinal);
        Assert.Contains($"sheet '{sheet}'", exception.Message, StringComparison.Ordinal);
        var after = RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Select(pivot => (pivot.SheetName, pivot.Name, pivot.CacheIndex))
            .ToArray();
        Assert.Equal(before, after);
        var retained = Assert.Single(after, pivot => pivot.Name == pivotName);
        Assert.Equal(sheet, retained.SheetName);
        Assert.Equal(original.CacheIndex, retained.CacheIndex);
        var retainedRead = RequireSuccess(_pivotCommands.Read(batch, pivotName, sheet));
        Assert.Equal(original.CacheIndex, retainedRead.PivotTable.CacheIndex);
    }

    [Fact]
    public void CreateFromDataModel_AllowsSameNameOnDifferentWorksheets()
    {
        var batch = _fixture.BatchToken;
        var firstSheet = _fixture.CreateTestSheet(batch);
        var secondSheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"RepeatedPivot{Guid.NewGuid():N}";

        RequireSuccess(_pivotCommands.CreateFromDataModel(
            batch, "SalesTable", firstSheet, "H1", pivotName));
        var secondResult = _pivotCommands.CreateFromDataModel(
            batch, "SalesTable", secondSheet, "H1", pivotName);

        RequireSuccess(secondResult);
        Assert.Equal(pivotName, secondResult.PivotTableName);
        var matchingPivots = RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Where(pivot => pivot.Name == pivotName)
            .ToArray();
        Assert.Equal(2, matchingPivots.Length);
        Assert.Contains(matchingPivots, pivot => pivot.SheetName == firstSheet);
        Assert.Contains(matchingPivots, pivot => pivot.SheetName == secondSheet);
        Assert.True(RequireSuccess(_pivotCommands.Read(batch, pivotName, firstSheet)).Success);
        Assert.True(RequireSuccess(_pivotCommands.Read(batch, pivotName, secondSheet)).Success);
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

    [Fact]
    public void Read_DataModelReportFilterReturnsExactSelectedMember()
    {
        Assert.True(
            PersistentServiceDataModelFixture.CreationResult.Success,
            "Data Model fixture must be created successfully");

        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        const string pivotName = "SemanticSnapshotFilterPivot";
        const string fieldName = "[CustomersTable].[Name]";
        const string memberName = "[CustomersTable].[Name].&[Acme Corp]";
        RequireSuccess(_pivotCommands.CreateFromDataModel(
            batch, "CustomersTable", sheet, "H1", pivotName));
        RequireSuccess(_pivotCommands.AddFilterField(batch, pivotName, fieldName));
        RequireSuccess(_pivotCommands.SetReportFilter(
            batch, sheet, pivotName, fieldName, [memberName]));

        var snapshot = RequireSuccess(_pivotCommands.Read(
            batch, pivotName, sheet, maxSelectedMembers: 1));

        Assert.True(snapshot.IsOlapAvailable);
        Assert.True(snapshot.IsOlap);
        Assert.True(snapshot.FieldSelectionsAvailable);
        Assert.True(snapshot.ValueFieldsAvailable);
        var selection = Assert.Single(snapshot.FieldSelections,
            item => item.FieldName == fieldName);
        Assert.Equal(PivotFieldArea.Filter, selection.Area);
        Assert.True(selection.LayoutAvailable);
        Assert.True(selection.SelectionStateAvailable);
        Assert.False(selection.ShowAll);
        Assert.Equal([memberName], selection.SelectedMemberUniqueNames);
        Assert.Equal(1, selection.SelectedMemberCount);
        Assert.False(selection.MembersTruncated);
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
                Assert.Equal(expected, Convert.ToDouble(((object)data.Value2),
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

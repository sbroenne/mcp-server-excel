using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Slicer")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class PersistentServiceOlapSlicerTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly ISlicerCommands _slicers = fixture.CreateCommands<ISlicerCommands>();
    private readonly IPivotTableCommands _pivots = fixture.CreateCommands<IPivotTableCommands>();
    private const string Sheet = "SlicerData";
    private const string Field = "[SlicerQuarters].[Quarter]";
    private static readonly string[] Captions = ["2026 Q1", "2026 Q2", "2026 Q3"];

    private SlicerResult CreateModelSlicer()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, Sheet);
        Assert.True(_commands.SetValues(batch, Sheet, "A1:B4",
            [["Key", "Quarter"], [1, Captions[0]], [2, Captions[1]], [3, Captions[2]]]).Success);
        Assert.True(_commands.SetValues(batch, Sheet, "D1:E5",
            [["QuarterKey", "Amount"], [1, 10], [2, 20], [2, 30], [3, 40]]).Success);
        var tables = _fixture.CreateCommands<ITableCommands>();
        Assert.True(tables.Create(batch, Sheet, "SlicerQuarters", "A1:B4").Success);
        _fixture.RegisterTableForCleanup("SlicerQuarters");
        Assert.True(tables.Create(batch, Sheet, "SlicerSales", "D1:E5").Success);
        _fixture.RegisterTableForCleanup("SlicerSales");
        Assert.True(tables.AddToDataModel(batch, "SlicerQuarters").Success);
        _fixture.RegisterDataModelTableForCleanup("SlicerQuarters");
        Assert.True(tables.AddToDataModel(batch, "SlicerSales").Success);
        _fixture.RegisterDataModelTableForCleanup("SlicerSales");
        var model = _fixture.CreateCommands<IDataModelCommands>();
        var relationships = _fixture.CreateCommands<IDataModelRelCommands>();
        Assert.True(relationships.CreateRelationship(batch, "SlicerSales", "QuarterKey",
            "SlicerQuarters", "Key").Success);
        Assert.True(model.CreateMeasure(batch, "SlicerSales", "SlicerTotal",
            "SUM(SlicerSales[Amount])").Success);
        _fixture.RegisterDataModelMeasureForCleanup("SlicerTotal");
        Assert.True(_pivots.CreateFromDataModel(batch, "SlicerSales", Sheet, "G1", "SlicerPivot").Success);
        var fields = _fixture.CreateCommands<IPivotTableFieldCommands>();
        Assert.True(fields.AddRowField(batch, "SlicerPivot", Field).Success);
        Assert.True(fields.AddValueField(batch, "SlicerPivot", "[Measures].[SlicerTotal]").Success);
        var created = _slicers.CreateSlicer(batch, "SlicerPivot", Field, "QuarterSlicer", Sheet, "K1");
        Assert.True(created.Success, created.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(created.ErrorMessage));
        return created;
    }

    [Fact]
    public void CreateAndList_DataModel_ReturnCaptionsAndReuseCache()
    {
        var created = CreateModelSlicer();
        Assert.Equal(Captions, created.AvailableItems.Order());
        Assert.Equal(Captions, created.SelectedItems.Order());
        Assert.Equal(["SlicerPivot"], created.ConnectedPivotTables);
        AssertSelection([Captions[1]], true, 50, Captions[1]);
        var second = _slicers.CreateSlicer(_fixture.BatchToken, "SlicerPivot", Field,
            "SecondQuarterSlicer", Sheet, "N1");
        Assert.True(second.Success, second.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(second.ErrorMessage));
        Assert.Equal("SecondQuarterSlicer", second.Name);
        Assert.Equal(Field, second.FieldName);
        Assert.Equal(Captions, second.AvailableItems.Order());
        Assert.Equal([Captions[1]], second.SelectedItems);
        Assert.Equal(["SlicerPivot"], second.ConnectedPivotTables);
        AssertPivot(50, Captions[1]);
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.SlicerCaches? caches = null;
            try
            {
                caches = ctx.Book.SlicerCaches;
                Assert.Equal(1, caches.Count);
            }
            finally { ComUtilities.Release(ref caches); }
        });
        var list = _slicers.ListSlicers(_fixture.BatchToken, "SlicerPivot");
        Assert.True(list.Success, list.ErrorMessage);
        Assert.Equal(2, list.Slicers.Count);
        Assert.All(list.Slicers, slicer =>
        {
            Assert.Equal(Field, slicer.FieldName);
            Assert.Equal(Captions, slicer.AvailableItems.Order());
            Assert.Equal([Captions[1]], slicer.SelectedItems);
            Assert.Equal(["SlicerPivot"], slicer.ConnectedPivotTables);
        });
        AssertPivot(50, Captions[1]);
    }

    [Fact]
    public void Selection_DataModel_OmittedServiceClearFirstReplacesExistingFilter()
    {
        CreateModelSlicer();
        AssertSelection([Captions[1]], true, 50, Captions[1]);
        var response = _fixture.Send("slicer.set-slicer-selection", new
        {
            slicerName = "QuarterSlicer",
            selectedItems = new[] { Captions[0] }
        });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
        Assert.NotNull(response.Result);
        using var result = JsonDocument.Parse(response.Result);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(Captions, result.RootElement.GetProperty("availableItems")
            .EnumerateArray().Select(item => item.GetString()).Order());
        Assert.Equal([Captions[0]], result.RootElement.GetProperty("selectedItems")
            .EnumerateArray().Select(item => item.GetString()));
        var listed = _slicers.ListSlicers(_fixture.BatchToken, "SlicerPivot");
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal([Captions[0]], Assert.Single(listed.Slicers).SelectedItems);
        AssertPivot(10, Captions[0]);
    }

    [Fact]
    public void Selection_DataModel_ReplaceAddClearAndUniqueNameChangePivotData()
    {
        CreateModelSlicer();
        AssertSelection([Captions[1]], true, 50, Captions[1]);
        AssertSelection([Captions[0]], false, 60, Captions[0], Captions[1]);
        AssertSelection(["2026 q1", Captions[0]], true, 10, Captions[0]);
        AssertSelection([Captions[2]], true, 40, Captions[2]);
        AssertSelection([], true, 100, Captions);
        AssertSelection([Captions[0]], false, 100, Captions);
        AssertSelection([Captions[1]], true, 50, Captions[1]);
        AssertSelection([], false, 100, Captions);
        var uniqueName = _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.SlicerCacheLevels? levels = null;
            Excel.SlicerCacheLevel? level = null;
            Excel.SlicerItems? items = null;
            Excel.SlicerItem? item = null;
            try
            {
                caches = ctx.Book.SlicerCaches;
                cache = caches.Item[1];
                levels = cache.SlicerCacheLevels;
                level = levels.Item[1];
                items = level.SlicerItems;
                item = items.Item[1];
                Assert.Equal(Captions[0], item.Caption);
                return item.Name;
            }
            finally
            {
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref level);
                ComUtilities.Release(ref levels);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
            }
        });
        AssertSelection([uniqueName], true, 10, Captions[0]);
    }

    [Theory]
    [InlineData(true, "missing")]
    [InlineData(false, "missing")]
    [InlineData(true, "")]
    [InlineData(false, "")]
    public void Selection_DataModel_InvalidItemFailsWithoutChangingFilter(bool clearFirst, string invalidItem)
    {
        CreateModelSlicer();
        AssertSelection([Captions[1]], true, 50, Captions[1]);
        var error = Assert.Throws<ArgumentException>(() =>
            _slicers.SetSlicerSelection(_fixture.BatchToken, "QuarterSlicer", [Captions[0], invalidItem], clearFirst));
        Assert.Contains("was not found", error.Message);
        var list = _slicers.ListSlicers(_fixture.BatchToken);
        Assert.True(list.Success, list.ErrorMessage);
        Assert.Equal([Captions[1]], Assert.Single(list.Slicers).SelectedItems);
        AssertPivot(50, Captions[1]);
        var missing = _slicers.SetSlicerSelection(_fixture.BatchToken, "MissingSlicer", [Captions[0]]);
        Assert.False(missing.Success);
        Assert.Contains("not found", missing.ErrorMessage);
        AssertSelection([Captions[2]], true, 40, Captions[2]);
    }

    [Fact]
    public async Task Selection_DataModel_SaveCloseReopenPreservesFilterAndRemainsEditable()
    {
        CreateModelSlicer();
        AssertSelection([Captions[1]], true, 50, Captions[1]);

        await _fixture.SaveAndReopenAsync();

        var listed = _slicers.ListSlicers(_fixture.BatchToken, "SlicerPivot");
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(listed.ErrorMessage));
        var slicer = Assert.Single(listed.Slicers);
        Assert.Equal("QuarterSlicer", slicer.Name);
        Assert.Equal(Field, slicer.FieldName);
        Assert.Equal(Captions, slicer.AvailableItems.Order());
        Assert.Equal([Captions[1]], slicer.SelectedItems);
        Assert.Equal(["SlicerPivot"], slicer.ConnectedPivotTables);
        AssertPivot(50, Captions[1]);

        AssertSelection([Captions[0]], false, 60, Captions[0], Captions[1]);
        AssertSelection([], false, 100, Captions);
        AssertSelection([Captions[2]], true, 40, Captions[2]);
    }

    [Fact]
    public void Selection_DataModel_SharedCacheFiltersBothConnectedPivots()
    {
        CreateModelSlicer();
        var batch = _fixture.BatchToken;
        Assert.True(_pivots.CreateFromDataModel(batch, "SlicerSales", Sheet, "G10",
            "SecondSlicerPivot").Success);
        var fields = _fixture.CreateCommands<IPivotTableFieldCommands>();
        Assert.True(fields.AddRowField(batch, "SecondSlicerPivot", Field).Success);
        Assert.True(fields.AddValueField(batch, "SecondSlicerPivot", "[Measures].[SlicerTotal]").Success);

        // Connection management has no public action; prepare the native shared-cache state.
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.SlicerPivotTables? connected = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[Sheet];
                pivots = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivots.Item("SecondSlicerPivot");
                caches = ctx.Book.SlicerCaches;
                cache = caches.Item[1];
                connected = cache.PivotTables;
                ct.ThrowIfCancellationRequested();
                connected.AddPivotTable(pivot);
                Assert.Equal(2, connected.Count);
            }
            finally
            {
                ComUtilities.Release(ref connected);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

        AssertPivot(100, Captions);
        AssertPivotAt(100, "G10:H15", Captions);
        var filtered = _slicers.SetSlicerSelection(batch, "QuarterSlicer", [Captions[1]]);
        Assert.True(filtered.Success, filtered.ErrorMessage);
        Assert.Equal([Captions[1]], filtered.SelectedItems);
        Assert.Equal(["SecondSlicerPivot", "SlicerPivot"], filtered.ConnectedPivotTables.Order());
        AssertPivot(50, Captions[1]);
        AssertPivotAt(50, "G10:H15", Captions[1]);

        var created = _slicers.CreateSlicer(batch, "SecondSlicerPivot", Field,
            "SecondQuarterSlicer", Sheet, "N10");
        Assert.True(created.Success, created.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(created.ErrorMessage));
        Assert.Equal(Captions, created.AvailableItems.Order());
        Assert.Equal([Captions[1]], created.SelectedItems);
        Assert.Equal(["SecondSlicerPivot", "SlicerPivot"], created.ConnectedPivotTables.Order());
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.SlicerCaches? caches = null;
            try
            {
                caches = ctx.Book.SlicerCaches;
                Assert.Equal(1, caches.Count);
            }
            finally { ComUtilities.Release(ref caches); }
        });
        AssertPivot(50, Captions[1]);
        AssertPivotAt(50, "G10:H15", Captions[1]);
        foreach (string pivotName in new[] { "SlicerPivot", "SecondSlicerPivot" })
        {
            var listed = _slicers.ListSlicers(batch, pivotName);
            Assert.True(listed.Success, listed.ErrorMessage);
            Assert.Equal(2, listed.Slicers.Count);
            Assert.All(listed.Slicers, slicer =>
            {
                Assert.Equal(Captions, slicer.AvailableItems.Order());
                Assert.Equal([Captions[1]], slicer.SelectedItems);
                Assert.Equal(["SecondSlicerPivot", "SlicerPivot"], slicer.ConnectedPivotTables.Order());
            });
        }

        var added = _slicers.SetSlicerSelection(batch, "SecondQuarterSlicer", [Captions[0]], false);
        Assert.True(added.Success, added.ErrorMessage);
        Assert.Equal([Captions[0], Captions[1]], added.SelectedItems.Order());
        AssertPivot(60, Captions[0], Captions[1]);
        AssertPivotAt(60, "G10:H15", Captions[0], Captions[1]);
        var cleared = _slicers.SetSlicerSelection(batch, "QuarterSlicer", [], false);
        Assert.True(cleared.Success, cleared.ErrorMessage);
        Assert.Equal(Captions, cleared.SelectedItems.Order());
        AssertPivot(100, Captions);
        AssertPivotAt(100, "G10:H15", Captions);
    }

    [Theory]
    [InlineData("MissingPivot", Field, typeof(InvalidOperationException))]
    [InlineData("SlicerPivot", "[SlicerQuarters].[Missing]", typeof(ArgumentException))]
    public void Create_DataModel_InvalidSourceFailsWithoutReturningEmptySuccess(string pivotName, string field, Type errorType)
    {
        CreateModelSlicer();
        AssertSelection([Captions[1]], true, 50, Captions[1]);
        Assert.Throws(errorType, () =>
            _slicers.CreateSlicer(_fixture.BatchToken, pivotName, field, "InvalidSlicer", Sheet, "N1"));
        var list = _slicers.ListSlicers(_fixture.BatchToken, "SlicerPivot");
        Assert.True(list.Success, list.ErrorMessage);
        var existing = Assert.Single(list.Slicers);
        Assert.Equal(Captions, existing.AvailableItems.Order());
        Assert.Equal([Captions[1]], existing.SelectedItems);
        AssertPivot(50, Captions[1]);
        AssertSelection([Captions[2]], true, 40, Captions[2]);
    }

    [Fact]
    public void NativeExcel_DataModel_LevelItemsAndVisibleListAreSupported()
    {
        CreateModelSlicer();
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.SlicerCacheLevels? levels = null;
            Excel.SlicerCacheLevel? level = null;
            Excel.SlicerItems? items = null;
            Excel.SlicerItem? item = null;
            try
            {
                caches = ctx.Book.SlicerCaches;
                cache = caches.Item[1];
                Assert.True(cache.OLAP);
                levels = cache.SlicerCacheLevels;
                level = levels.Item[1];
                items = level.SlicerItems;
                Assert.Equal(3, items.Count);
                var captions = new List<string>();
                string? selectedName = null;
                for (int i = 1; i <= items.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        item = items.Item[i];
                        captions.Add(item.Caption);
                        Assert.True(item.Selected);
                        if (item.Caption == Captions[1])
                            selectedName = item.Name;
                    }
                    finally { ComUtilities.Release(ref item); }
                }
                Assert.Equal(Captions, captions.Order());
                Assert.NotNull(selectedName);
                cache.VisibleSlicerItemsList = new[] { selectedName };
                Array names = Assert.IsAssignableFrom<Array>((object)cache.VisibleSlicerItemsList);
                Assert.Equal(new[] { selectedName }, names.Cast<string>());
            }
            finally
            {
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref level);
                ComUtilities.Release(ref levels);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
            }
        });
        AssertPivot(50, Captions[1]);
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            try
            {
                caches = ctx.Book.SlicerCaches;
                cache = caches.Item[1];
                cache.ClearManualFilter();
            }
            finally
            {
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
            }
        });
        AssertPivot(100, Captions);
    }

    private void AssertSelection(List<string> requested, bool clearFirst, double total, params string[] captions)
    {
        var result = _slicers.SetSlicerSelection(_fixture.BatchToken, "QuarterSlicer", requested, clearFirst);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(Captions, result.AvailableItems.Order());
        Assert.Equal(captions.Order(), result.SelectedItems.Order());
        var listed = _slicers.ListSlicers(_fixture.BatchToken, "SlicerPivot");
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(captions.Order(), Assert.Single(listed.Slicers).SelectedItems.Order());
        AssertPivot(total, captions);
    }

    private void AssertPivot(double total, params string[] captions)
        => AssertPivotAt(total, "G1:H6", captions);

    private void AssertPivotAt(double total, string address, params string[] captions)
    {
        var values = _commands.GetValues(_fixture.BatchToken, Sheet, address);
        Assert.True(values.Success, values.ErrorMessage);
        var rows = values.Values.Where(row => row[0] is not null
            && double.TryParse(row[1]?.ToString(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out _)).ToList();
        Assert.Equal(captions.Length + 1, rows.Count);
        Assert.Equal(captions.Order(), rows.Take(captions.Length).Select(row => row[0]!.ToString()).Order());
        foreach (var row in rows.Take(captions.Length))
        {
            Assert.Equal(row[0]!.ToString() switch
            {
                "2026 Q1" => 10d,
                "2026 Q2" => 50d,
                "2026 Q3" => 40d,
                _ => throw new InvalidOperationException("Unexpected PivotTable caption.")
            }, Convert.ToDouble(row[1], System.Globalization.CultureInfo.InvariantCulture));
        }
        Assert.Equal(total, double.Parse(rows[^1][1]!.ToString()!, System.Globalization.CultureInfo.InvariantCulture));
    }
}

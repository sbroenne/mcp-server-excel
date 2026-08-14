using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

public partial class PivotTableCommandsTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void CacheOptions_SetAndGet_RoundTripsRegularPivotCacheSettings()
    {
        var testFile = CreateTestFileWithData(nameof(CacheOptions_SetAndGet_RoundTripsRegularPivotCacheSettings));
        using var batch = ExcelSession.BeginBatch(testFile);

        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "CacheOptionsPivot");
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var setResult = _pivotCommands.SetCacheOptions(
            batch,
            "CacheOptionsPivot",
            refreshOnFileOpen: true,
            missingItemsLimit: PivotMissingItemsLimit.None,
            saveSourceData: false);

        Assert.True(setResult.Success, setResult.ErrorMessage);
        var getResult = _pivotCommands.GetCacheOptions(batch, "CacheOptionsPivot");
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.True(getResult.RefreshOnFileOpen);
        Assert.Equal(PivotMissingItemsLimit.None, getResult.MissingItemsLimit);
        Assert.False(getResult.SaveSourceData);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public void CacheOptions_OlapCache_RejectsUnsupportedMutations()
    {
        Assert.True(_creationResult.Success, _creationResult.ErrorMessage);
        bool optimizeCache;
        using (var readBatch = ExcelSession.BeginBatch(_pivotFile))
        {
            var current = _pivotCommands.GetCacheOptions(readBatch, "DataModelPivot");
            Assert.True(current.Success, current.ErrorMessage);
            Assert.True(current.IsOlap);
            optimizeCache = current.OptimizeCache;
        }

        using (var missingItemsBatch = ExcelSession.BeginBatch(_pivotFile))
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                _pivotCommands.SetCacheOptions(
                    missingItemsBatch,
                    "DataModelPivot",
                    missingItemsLimit: PivotMissingItemsLimit.None));
            Assert.Contains("not available for OLAP", exception.Message);
        }

        using (var optimizeBatch = ExcelSession.BeginBatch(_pivotFile))
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                _pivotCommands.SetCacheOptions(
                    optimizeBatch,
                    "DataModelPivot",
                    optimizeCache: !optimizeCache));
            Assert.Contains("read-only for external OLE DB/OLAP", exception.Message);
        }

        using (var saveSourceDataBatch = ExcelSession.BeginBatch(_pivotFile))
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                _pivotCommands.SetCacheOptions(
                    saveSourceDataBatch,
                    "DataModelPivot",
                    saveSourceData: true));
            Assert.Contains("cannot save source records", exception.Message);
        }
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupItems_ThenUngroupField_RestoresOriginalFieldLayout()
    {
        var testFile = CreateTestFileWithData(nameof(GroupItems_ThenUngroupField_RestoresOriginalFieldLayout));
        using var batch = ExcelSession.BeginBatch(testFile);
        batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? sourceRows = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["SalesData"];
                sourceRows = sheet.Range["A7:D9"];
                sourceRows.Value2 = new object[,]
                {
                    { "West", "Widget", 175, new DateTime(2025, 3, 10) },
                    { "Group Existing", "Widget", 225, new DateTime(2025, 3, 15) },
                    { "East", "Gadget", 250, new DateTime(2025, 3, 20) }
                };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sourceRows);
                ComUtilities.Release(ref sheet);
            }
        });

        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D9", "SalesData", "F2", "ManualGroupingPivot");
        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.True(_pivotCommands.AddRowField(batch, "ManualGroupingPivot", "Region").Success);
        Assert.True(_pivotCommands.AddValueField(batch, "ManualGroupingPivot", "Sales").Success);

        var groupResult = _pivotCommands.GroupItems(
            batch,
            "ManualGroupingPivot",
            "Region",
            ["North", "South"],
            "All Regions");

        Assert.True(groupResult.Success, groupResult.ErrorMessage);
        Assert.Equal("All Regions", groupResult.GroupName);
        Assert.Equal(["North", "South"], groupResult.Items);
        Assert.False(string.IsNullOrWhiteSpace(groupResult.GroupedFieldName));

        var groupedFields = _pivotCommands.ListFields(batch, "ManualGroupingPivot");
        Assert.Contains(groupedFields.Fields, field => field.Name == groupResult.GroupedFieldName);
        var firstGroupItems = ReadPivotItemNames(batch, "ManualGroupingPivot", groupResult.GroupedFieldName);
        Assert.Contains("All Regions", firstGroupItems);
        Assert.Contains("Group Existing", firstGroupItems);

        var secondGroupResult = _pivotCommands.GroupItems(
            batch,
            "ManualGroupingPivot",
            "Region",
            ["West", "East"],
            "Outer Regions");
        Assert.True(secondGroupResult.Success, secondGroupResult.ErrorMessage);
        Assert.Equal(groupResult.GroupedFieldName, secondGroupResult.GroupedFieldName);

        var repeatedGroupItems = ReadPivotItemNames(batch, "ManualGroupingPivot", groupResult.GroupedFieldName);
        Assert.Contains("All Regions", repeatedGroupItems);
        Assert.Contains("Outer Regions", repeatedGroupItems);
        Assert.Contains("Group Existing", repeatedGroupItems);

        var ungroupResult = _pivotCommands.UngroupField(
            batch,
            "ManualGroupingPivot",
            groupResult.GroupedFieldName);
        Assert.True(ungroupResult.Success, ungroupResult.ErrorMessage);

        var restoredFields = _pivotCommands.ListFields(batch, "ManualGroupingPivot");
        Assert.DoesNotContain(restoredFields.Fields, field => field.Name == groupResult.GroupedFieldName);
        Assert.Contains(restoredFields.Fields, field => field.Name == "Region");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DrillThrough_DataCell_CreatesDetailWorksheetWithSourceRows()
    {
        var testFile = CreateTestFileWithData(nameof(DrillThrough_DataCell_CreatesDetailWorksheetWithSourceRows));
        using var batch = ExcelSession.BeginBatch(testFile);

        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "DrillThroughPivot");
        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.True(_pivotCommands.AddRowField(batch, "DrillThroughPivot", "Region").Success);
        Assert.True(_pivotCommands.AddValueField(batch, "DrillThroughPivot", "Sales").Success);

        var dataCellAddress = batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivotTables = null;
            Excel.PivotTable? pivot = null;
            Excel.Range? dataBodyRange = null;
            Excel.Range? firstDataCell = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["SalesData"];
                pivotTables = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivotTables.Item("DrillThroughPivot");
                dataBodyRange = pivot.DataBodyRange;
                firstDataCell = (Excel.Range)dataBodyRange.Cells[1, 1];
                return firstDataCell.Address;
            }
            finally
            {
                ComUtilities.Release(ref firstDataCell);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivotTables);
                ComUtilities.Release(ref sheet);
            }
        });

        var result = _pivotCommands.DrillThrough(batch, "DrillThroughPivot", dataCellAddress);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(result.DetailSheetName));
        Assert.True(result.DetailRowCount > 1);

        var detailExists = batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? detailSheet = null;
            try
            {
                detailSheet = (Excel.Worksheet)ctx.Book.Worksheets[result.DetailSheetName];
                return detailSheet.Name == result.DetailSheetName;
            }
            finally
            {
                ComUtilities.Release(ref detailSheet);
            }
        });
        Assert.True(detailExists);
    }

    private static HashSet<string> ReadPivotItemNames(
        IExcelBatch batch,
        string pivotTableName,
        string fieldName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivotTables = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotField? field = null;
            Excel.PivotItems? items = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["SalesData"];
                pivotTables = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivotTables.Item(pivotTableName);
                field = (Excel.PivotField)pivot.PivotFields(fieldName);
                items = (Excel.PivotItems)field.PivotItems();

                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var index = 1; index <= items.Count; index++)
                {
                    Excel.PivotItem? item = null;
                    try
                    {
                        item = items.Item(index);
                        names.Add(item.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }

                return names;
            }
            finally
            {
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivotTables);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

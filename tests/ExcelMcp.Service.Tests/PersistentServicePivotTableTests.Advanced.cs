using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void GroupItems_ThenUngroupField_RestoresOriginalFieldLayout()
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? sourceRows = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[_salesSheetName];
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
            batch, _salesSheetName, "A1:D9", destinationSheet, "A1", "ManualGroupingPivot");
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
        var firstGroupItems = ReadPivotItemNames("ManualGroupingPivot", groupResult.GroupedFieldName, destinationSheet);
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

        var repeatedGroupItems = ReadPivotItemNames("ManualGroupingPivot", groupResult.GroupedFieldName, destinationSheet);
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
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", destinationSheet, "A1", "DrillThroughPivot");
        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.True(_pivotCommands.AddRowField(batch, "DrillThroughPivot", "Region").Success);
        Assert.True(_pivotCommands.AddValueField(batch, "DrillThroughPivot", "Sales").Success);

        var dataCellAddress = _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivotTables = null;
            Excel.PivotTable? pivot = null;
            Excel.Range? dataBodyRange = null;
            Excel.Range? firstDataCell = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[destinationSheet];
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
        _fixture.RegisterSheetForCleanup(result.DetailSheetName);

        var detailExists = _fixture.ExecuteRawVerification((ctx, ct) =>
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

    private HashSet<string> ReadPivotItemNames(
        string pivotTableName,
        string fieldName,
        string sheetName)
    {
        return _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivotTables = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotField? field = null;
            Excel.PivotItems? items = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[sheetName];
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

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
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "A7:D9",
            [
                ["West", "Widget", 175, "2025-03-10"],
                ["Group Existing", "Widget", 225, "2025-03-15"],
                ["East", "Gadget", 250, "2025-03-20"]
            ]));

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D9", destinationSheet, "A1", "ManualGroupingPivot");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "ManualGroupingPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "ManualGroupingPivot", "Sales"));

        var groupResult = _pivotCommands.GroupItems(
            batch,
            "ManualGroupingPivot",
            "Region",
            ["North", "South"],
            "All Regions");

        RequireSuccess(groupResult);
        Assert.Equal("All Regions", groupResult.GroupName);
        Assert.Equal(["North", "South"], groupResult.Items);
        Assert.False(string.IsNullOrWhiteSpace(groupResult.GroupedFieldName));

        var groupedFields = _pivotCommands.ListFields(batch, "ManualGroupingPivot");
        RequireSuccess(groupedFields);
        Assert.Contains(groupedFields.Fields, field => field.Name == groupResult.GroupedFieldName);
        var firstGroupItems = ReadPivotItemNames("ManualGroupingPivot", groupResult.GroupedFieldName, destinationSheet);
        Assert.Contains("All Regions", firstGroupItems);
        Assert.Contains("Group Existing", firstGroupItems);
        var firstGroupedData = RequireSuccess(_pivotCommands.GetData(batch, "ManualGroupingPivot"));
        var allRegions = Assert.Single(firstGroupedData.Values, row => row[0]?.ToString() == "All Regions");
        Assert.Equal(650d, Convert.ToDouble(allRegions[^1], System.Globalization.CultureInfo.InvariantCulture));

        var secondGroupResult = _pivotCommands.GroupItems(
            batch,
            "ManualGroupingPivot",
            "Region",
            ["West", "East"],
            "Outer Regions");
        RequireSuccess(secondGroupResult);
        Assert.Equal(groupResult.GroupedFieldName, secondGroupResult.GroupedFieldName);

        var repeatedGroupItems = ReadPivotItemNames("ManualGroupingPivot", groupResult.GroupedFieldName, destinationSheet);
        Assert.Contains("All Regions", repeatedGroupItems);
        Assert.Contains("Outer Regions", repeatedGroupItems);
        Assert.Contains("Group Existing", repeatedGroupItems);
        var repeatedData = RequireSuccess(_pivotCommands.GetData(batch, "ManualGroupingPivot"));
        var outerRegions = Assert.Single(repeatedData.Values, row => row[0]?.ToString() == "Outer Regions");
        Assert.Equal(425d, Convert.ToDouble(outerRegions[^1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(650d, Convert.ToDouble(
            Assert.Single(repeatedData.Values, row => row[0]?.ToString() == "All Regions")[^1],
            System.Globalization.CultureInfo.InvariantCulture));

        var ungroupResult = _pivotCommands.UngroupField(
            batch,
            "ManualGroupingPivot",
            groupResult.GroupedFieldName);
        RequireSuccess(ungroupResult);

        var restoredFields = _pivotCommands.ListFields(batch, "ManualGroupingPivot");
        RequireSuccess(restoredFields);
        Assert.DoesNotContain(restoredFields.Fields, field => field.Name == groupResult.GroupedFieldName);
        Assert.Contains(restoredFields.Fields, field => field.Name == "Region");
        Assert.Equal(["East", "Group Existing", "North", "South", "West"],
            ReadPivotItemNames("ManualGroupingPivot", "Region", destinationSheet).Order(StringComparer.Ordinal));
        AssertOriginalSales();
        var data = RequireSuccess(_pivotCommands.GetData(batch, "ManualGroupingPivot"));
        Assert.Equal(1300d, Convert.ToDouble(data.Values[^1][^1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DrillThrough_DataCell_CreatesDetailWorksheetWithSourceRows()
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", destinationSheet, "A1", "DrillThroughPivot");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "DrillThroughPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "DrillThroughPivot", "Sales"));

        var dataCellAddress = _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivotTables = null;
            Excel.PivotTable? pivot = null;
            Excel.Range? dataBodyRange = null;
            Excel.Range? firstDataCell = null;
            Excel.Range? cells = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, destinationSheet);
                pivotTables = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivotTables.Item("DrillThroughPivot");
                dataBodyRange = pivot.DataBodyRange;
                cells = dataBodyRange.Cells;
                firstDataCell = (Excel.Range)cells[1, 1];
                return firstDataCell.Address;
            }
            finally
            {
                ComUtilities.Release(ref firstDataCell);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivotTables);
                ComUtilities.Release(ref sheet);
            }
        });

        var result = _pivotCommands.DrillThrough(batch, "DrillThroughPivot", dataCellAddress);

        RequireSuccess(result);
        Assert.False(string.IsNullOrWhiteSpace(result.DetailSheetName));
        RequireSuccess(result);
        Assert.Equal(4, result.DetailRowCount);
        _fixture.RegisterSheetForCleanup(result.DetailSheetName);

        var detailExists = _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? detailSheet = null;
            try
            {
                detailSheet = ComUtilities.FindSheet(ctx.Book, result.DetailSheetName);
                return detailSheet.Name == result.DetailSheetName;
            }
            finally
            {
                ComUtilities.Release(ref detailSheet);
            }
        });
        Assert.True(detailExists);
        var details = RequireSuccess(_commands.GetValues(batch, result.DetailSheetName, "A1:D4")).Values;
        Assert.Equal(4, details.Count);
        Assert.Equal(new object?[] { "Region", "Product", "Sales", "Date" }, details[0]);
        var records = details.Skip(1).OrderBy(row =>
            Convert.ToDouble(row[3], System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        var sales = new[] { 100d, 150d, 75d };
        var products = new[] { "Widget", "Widget", "Gadget" };
        var dates = new[] { new DateTime(2025, 1, 15), new DateTime(2025, 1, 20), new DateTime(2025, 2, 15) };
        for (var index = 0; index < records.Length; index++)
        {
            var row = records[index];
            Assert.Equal(4, row.Count);
            Assert.Equal("North", row[0]);
            Assert.Equal(products[index], row[1]);
            Assert.Equal(sales[index], Convert.ToDouble(row[2], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(dates[index], DateTime.FromOADate(Convert.ToDouble(row[3],
                System.Globalization.CultureInfo.InvariantCulture)));
        }
        AssertOriginalSales();
        AssertPivotSales(325, 325, "DrillThroughPivot");
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
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
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

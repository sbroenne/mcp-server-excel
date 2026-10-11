using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public async Task ExternalOlapPivotTable_ReportFilter_SelectsOneMultipleAndAllWithoutChangingOtherReport()
    {
        var (sheet, pivot, _) = CreateConfiguredExternalPivot("ExternalReportFilter");
        var (otherSheet, otherPivot, _) = CreateConfiguredExternalPivot("ExternalReportFilterOther");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var source = RequireSuccess(pivots.GetConnection(batch, sheet, pivot));
        Assert.True(source.IsExternal);
        Assert.True(source.IsOlap);
        Assert.False(source.IsDataModel);
        RequireSuccess(pivots.AddRowField(batch, pivot, ExternalRelativePeriod));
        RequireSuccess(pivots.AddValueField(batch, pivot, ExternalAmountMeasure));
        RequireSuccess(pivots.AddRowField(batch, otherPivot, ExternalCalendar));
        RequireSuccess(pivots.AddValueField(batch, otherPivot, ExternalAmountMeasure));
        AssertExternalGrandTotal(pivot, 300);
        AssertExternalGrandTotal(otherPivot, 300);
        var otherBefore = ReadExternalPivotState(otherSheet, otherPivot);
        var members = ReadExternalReportFilterMembers(sheet, pivot);
        Assert.Equal(["Current Month", "Earlier Months", "Next Month", "Prior Month"],
            members.Keys.Order(StringComparer.Ordinal));
        RequireSuccess(pivots.RemoveField(batch, pivot, ExternalRelativePeriod));
        RequireSuccess(pivots.AddFilterField(batch, pivot, ExternalRelativePeriod));

        Select([members["Current Month"]], 23);
        Select([members["Current Month"], members["Prior Month"]], 45);
        Select([members["Next Month"]], 24);
        Select([], 300);
        Select([members["Current Month"]], 23);
        Select([], 300);

        var rejected = pivots.SetReportFilter(batch, sheet, pivot, ExternalRelativePeriod,
            ["[Sales].[RelativePeriod].[RelativePeriod].&[MissingPeriod]"]);
        Assert.False(rejected.Success);
        Assert.False(string.IsNullOrWhiteSpace(rejected.ErrorMessage));
        Assert.True(rejected.MayHavePartiallyChanged);
        Assert.False(rejected.RollbackAttempted);
        Assert.True(rejected.ShowAll);
        Assert.Empty(rejected.SelectedItems);
        AssertTotal(300);
        Assert.Equal(otherBefore, ReadExternalPivotState(otherSheet, otherPivot));

        Select([members["Current Month"], members["Prior Month"]], 45);
        await _fixture.SaveAndReopenAsync();
        AssertTotal(45);
        Assert.Equal(otherBefore, ReadExternalPivotState(otherSheet, otherPivot));
        RequireSuccess(pivots.Refresh(batch, pivot));
        AssertTotal(45);
        Select([members["Current Month"], members["Prior Month"]], 45);
        Select([], 300);

        void Select(List<string> selectedItems, double expected)
        {
            var result = RequireSuccess(pivots.SetReportFilter(
                batch, sheet, pivot, ExternalRelativePeriod, selectedItems));
            Assert.Equal(sheet, result.SheetName);
            Assert.Equal(pivot, result.PivotTableName);
            Assert.Equal(ExternalRelativePeriod, result.FieldName);
            Assert.Equal(selectedItems.Count == 0, result.ShowAll);
            Assert.Equal(selectedItems.Order(StringComparer.Ordinal),
                result.SelectedItems.Order(StringComparer.Ordinal));
            Assert.False(result.MayHavePartiallyChanged);
            Assert.False(result.RollbackAttempted);
            AssertTotal(expected);
            Assert.Equal(otherBefore, ReadExternalPivotState(otherSheet, otherPivot));
        }

        void AssertTotal(double expected)
        {
            var data = RequireSuccess(pivots.GetData(batch, pivot));
            Assert.Equal(expected, Convert.ToDouble(
                Assert.Single(data.Values.SelectMany(row => row), value => value is double or int),
                CultureInfo.InvariantCulture));
        }
    }

    private Dictionary<string, string> ReadExternalReportFilterMembers(string sheetName, string pivotName) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubes = null;
            Excel.CubeField? cube = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? field = null;
            Excel.PivotItems? items = null;
            Excel.PivotItem? item = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                pivots = sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubes = pivot.CubeFields;
                cube = cubes[ExternalRelativePeriod];
                fields = cube.PivotFields;
                field = fields.Item(1);
                items = field.PivotItems();
                var members = new Dictionary<string, string>(StringComparer.Ordinal);
                for (int index = 1; index <= items.Count; index++)
                {
                    item = items.Item(index);
                    members.Add(Convert.ToString(item.Caption, CultureInfo.InvariantCulture)!, item.Name);
                    ComUtilities.Release(ref item);
                }
                return members;
            }
            finally
            {
                ComUtilities.Release(ref item);
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref cube);
                ComUtilities.Release(ref cubes);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
            }
        });
}

using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
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

        var wrongIdentifier = pivots.SetReportFilter(
            batch, sheet, pivot, $"{ExternalRelativePeriod}.[RelativePeriod]", []);
        Assert.False(wrongIdentifier.Success);
        Assert.Contains("list-fields", wrongIdentifier.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.False(wrongIdentifier.MayHavePartiallyChanged);

        SetExternalReportFilterSingleSelection(sheet, pivot, members["Current Month"]);
        var singleModeBefore = RequireSuccess(pivots.GetReportFilter(
            batch, sheet, pivot, ExternalRelativePeriod));
        Assert.False(singleModeBefore.MultipleSelectionEnabled);
        Select([members["Prior Month"]], 22);
        var singleModeAfter = RequireSuccess(pivots.GetReportFilter(
            batch, sheet, pivot, ExternalRelativePeriod));
        Assert.False(singleModeAfter.MultipleSelectionEnabled);

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
        Assert.Equal(["[Sales].[RelativePeriod].[RelativePeriod].&[MissingPeriod]"],
            rejected.RequestedItems);
        Assert.True(rejected.SelectionStateAvailable);
        Assert.Equal("assign-selected-items", rejected.FailurePhase);
        Assert.True(rejected.MayHavePartiallyChanged);
        Assert.False(rejected.RollbackAttempted);
        Assert.True(rejected.ShowAll);
        Assert.Empty(rejected.SelectedItems);
        var afterFailure = RequireSuccess(pivots.GetReportFilter(
            batch, sheet, pivot, ExternalRelativePeriod));
        Assert.True(afterFailure.ShowAll);
        Assert.True(afterFailure.SelectionStateAvailable);
        Assert.Empty(afterFailure.SelectedItems);
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
            var before = RequireSuccess(pivots.GetReportFilter(
                batch, sheet, pivot, ExternalRelativePeriod));
            var result = RequireSuccess(pivots.SetReportFilter(
                batch, sheet, pivot, ExternalRelativePeriod, selectedItems));
            Assert.Equal(sheet, result.SheetName);
            Assert.Equal(pivot, result.PivotTableName);
            Assert.Equal(ExternalRelativePeriod, result.FieldName);
            Assert.NotNull(result.RequestedItems);
            Assert.Equal(selectedItems.Order(StringComparer.Ordinal),
                result.RequestedItems!.Order(StringComparer.Ordinal));
            Assert.True(result.SelectionStateAvailable);
            Assert.Equal(selectedItems.Count == 0, result.ShowAll);
            Assert.Equal(selectedItems.Order(StringComparer.Ordinal),
                result.SelectedItems.Order(StringComparer.Ordinal));
            Assert.Equal(before.MultipleSelectionEnabled || selectedItems.Count > 1,
                result.MultipleSelectionEnabled);
            Assert.False(result.MayHavePartiallyChanged);
            Assert.False(result.RollbackAttempted);

            var inspected = RequireSuccess(pivots.GetReportFilter(
                batch, sheet, pivot, ExternalRelativePeriod));
            Assert.Equal(PivotFieldArea.Filter, inspected.FieldArea);
            Assert.Equal(result.ShowAll, inspected.ShowAll);
            Assert.Equal(result.MultipleSelectionEnabled, inspected.MultipleSelectionEnabled);
            Assert.Equal(result.SelectedItems.Order(StringComparer.Ordinal),
                inspected.SelectedItems.Order(StringComparer.Ordinal));
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

    private void SetExternalReportFilterSingleSelection(
        string sheetName, string pivotName, string selectedMember)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubes = null;
            Excel.CubeField? cube = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? field = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                pivots = sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubes = pivot.CubeFields;
                cube = cubes[ExternalRelativePeriod];
                fields = cube.PivotFields;
                field = fields.Item(1);
                cube.EnableMultiplePageItems = false;
                field.CurrentPageName = selectedMember;
                return true;
            }
            finally
            {
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

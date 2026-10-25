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
    public async Task ExternalOlapPivotTable_ReportFilter_ProviderAllPreservesSingleSelection()
    {
        var (sheet, pivot, _) = CreateConfiguredExternalPivot("ExternalNonstandardAll");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddRowField(batch, pivot, ExternalRelativePeriod));
        RequireSuccess(pivots.AddValueField(batch, pivot, ExternalAmountMeasure));
        AssertExternalGrandTotal(pivot, 300);
        var members = ReadExternalPivotMembers(sheet, pivot, ExternalRelativePeriod);
        string selectedMember = members["Current Month"];
        RequireSuccess(pivots.RemoveField(batch, pivot, ExternalRelativePeriod));
        RequireSuccess(pivots.AddFilterField(batch, pivot, ExternalRelativePeriod));
        string allMember = AssertNativeSelection(true, null);
        if (string.Equals(Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_SYNTHETIC_FIXTURE"), "true", StringComparison.OrdinalIgnoreCase))
        {
            Assert.Equal("[Sales].[RelativePeriod].[ALL].[AllMember]", allMember);
            Assert.Equal("[Sales].[RelativePeriod].[ALL].[AllMember].[Current Month]", selectedMember);
        }

        SetExternalReportFilterSingleSelection(sheet, pivot, selectedMember);
        AssertNativeSelection(false, selectedMember);
        AssertSelection(false, selectedMember, 23);

        SetExternalReportFilterSingleSelection(sheet, pivot, allMember);
        AssertNativeSelection(true, allMember);
        AssertSelection(true, null, 300);

        var selected = RequireSuccess(pivots.SetReportFilter(
            batch, sheet, pivot, ExternalRelativePeriod, [selectedMember]));
        Assert.False(selected.ShowAll);
        Assert.False(selected.MultipleSelectionEnabled);
        Assert.Equal([selectedMember], selected.SelectedItems);
        AssertNativeSelection(false, selectedMember);
        AssertSelection(false, selectedMember, 23);

        var assignedAll = RequireSuccess(pivots.SetReportFilter(
            batch, sheet, pivot, ExternalRelativePeriod, [allMember]));
        Assert.Equal([allMember], assignedAll.RequestedItems);
        Assert.True(assignedAll.SelectionStateAvailable);
        Assert.True(assignedAll.ShowAll);
        Assert.False(assignedAll.MultipleSelectionEnabled);
        Assert.Empty(assignedAll.SelectedItems);
        AssertSelection(true, null, 300);

        var rejected = pivots.SetReportFilter(batch, sheet, pivot, ExternalRelativePeriod,
            ["[Sales].[RelativePeriod].[ALL].[AllMember].[MissingPeriod]"]);
        Assert.False(rejected.Success);
        Assert.Equal("set-single-selection", rejected.FailurePhase);
        Assert.True(rejected.MayHavePartiallyChanged);
        Assert.False(rejected.RollbackAttempted);
        Assert.False(string.IsNullOrWhiteSpace(rejected.ErrorMessage));
        Assert.True(rejected.SelectionStateAvailable);
        Assert.True(rejected.ShowAll);
        Assert.Empty(rejected.SelectedItems);
        AssertSelection(true, null, 300);

        RequireSuccess(pivots.SetReportFilter(
            batch, sheet, pivot, ExternalRelativePeriod, [selectedMember]));
        var cleared = RequireSuccess(pivots.SetReportFilter(
            batch, sheet, pivot, ExternalRelativePeriod, []));
        Assert.True(cleared.SelectionStateAvailable);
        Assert.True(cleared.ShowAll);
        Assert.False(cleared.MultipleSelectionEnabled);
        Assert.Empty(cleared.SelectedItems);
        AssertNativeSelection(true, allMember);
        AssertSelection(true, null, 300);

        await _fixture.SaveAndReopenAsync();
        var disconnected = pivots.GetReportFilter(batch, sheet, pivot, ExternalRelativePeriod);
        Assert.False(disconnected.Success);
        Assert.False(disconnected.SelectionStateAvailable);
        Assert.Equal("read-filter-state", disconnected.FailurePhase);
        Assert.Contains("ADO session", disconnected.ErrorMessage, StringComparison.Ordinal);
        Assert.False(disconnected.MayHavePartiallyChanged);
        var disconnectedSnapshot = RequireSuccess(pivots.Read(batch, pivot, sheet));
        Assert.False(disconnectedSnapshot.SnapshotComplete);
        var unavailable = Assert.Single(disconnectedSnapshot.FieldSelections);
        Assert.False(unavailable.SelectionStateAvailable);
        Assert.Null(unavailable.ShowAll);
        Assert.Null(unavailable.SelectedMemberCount);
        Assert.Empty(unavailable.SelectedMemberUniqueNames);
        Assert.Contains("ADO session", unavailable.SelectionError, StringComparison.Ordinal);
        RequireSuccess(pivots.Refresh(batch, pivot));
        AssertNativeSelection(true, allMember);
        AssertSelection(true, null, 300);

        void AssertSelection(bool showAll, string? member, double total)
        {
            var snapshot = RequireSuccess(pivots.Read(batch, pivot, sheet));
            Assert.True(snapshot.SnapshotComplete);
            var field = Assert.Single(snapshot.FieldSelections);
            Assert.Equal(ExternalRelativePeriod, field.FieldName);
            Assert.Equal(PivotFieldArea.Filter, field.Area);
            Assert.True(field.SelectionStateAvailable);
            Assert.Equal(showAll, field.ShowAll);
            Assert.False(field.MultipleSelectionEnabled);
            Assert.Equal(showAll ? null : (int?)1, field.SelectedMemberCount);
            Assert.Equal(member is null ? [] : new[] { member }, field.SelectedMemberUniqueNames);
            Assert.False(field.MembersTruncated);

            var filter = RequireSuccess(pivots.GetReportFilter(batch, sheet, pivot, ExternalRelativePeriod));
            Assert.True(filter.SelectionStateAvailable);
            Assert.Equal(showAll, filter.ShowAll);
            Assert.False(filter.MultipleSelectionEnabled);
            Assert.Equal(member is null ? [] : new[] { member }, filter.SelectedItems);
            var data = RequireSuccess(pivots.GetData(batch, pivot));
            Assert.Equal(total, Convert.ToDouble(
                Assert.Single(data.Values.SelectMany(row => row), value => value is double or int),
                CultureInfo.InvariantCulture));
        }

        string AssertNativeSelection(bool showAll, string? currentPage) =>
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? worksheet = null;
                Excel.PivotTables? tables = null;
                Excel.PivotTable? table = null;
                Excel.CubeFields? cubes = null;
                Excel.CubeField? cube = null;
                Excel.PivotFields? fields = null;
                Excel.PivotField? field = null;
                Excel.PivotCache? cache = null;
                ADODB._Connection? ado = null;
                ADODB._Recordset? rows = null;
                ADODB.Fields? columns = null;
                ADODB.Field? all = null;
                try
                {
                    worksheet = ComUtilities.FindSheet(context.Book, sheet);
                    Assert.NotNull(worksheet);
                    tables = worksheet.PivotTables();
                    table = tables.Item(pivot);
                    cubes = table.CubeFields;
                    cube = cubes[ExternalRelativePeriod];
                    fields = cube.PivotFields;
                    field = fields.Item(1);
                    Assert.False(cube.EnableMultiplePageItems);
                    if (currentPage is not null)
                        Assert.Equal(currentPage, field.CurrentPageName);
                    cache = table.PivotCache();
                    ado = (ADODB._Connection)cache.ADOConnection;
                    rows = ado.OpenSchema((ADODB.SchemaEnum)34, new object?[]
                        { null, null, Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_CUBE")!, null, null, ExternalRelativePeriod }, Type.Missing);
                    Assert.False(Convert.ToBoolean(rows.EOF, CultureInfo.InvariantCulture));
                    columns = rows.Fields;
                    all = columns["ALL_MEMBER"];
                    string nativeAll = Convert.ToString((object)all.Value, CultureInfo.InvariantCulture)!;
                    Assert.False(string.IsNullOrWhiteSpace(nativeAll));
                    Assert.Equal(showAll, string.Equals(field.CurrentPageName, nativeAll, StringComparison.Ordinal));
                    return nativeAll;
                }
                finally
                {
                    ComUtilities.Release(ref all);
                    ComUtilities.Release(ref columns);
                    try
                    {
                        if (rows is not null)
                            rows.Close();
                    }
                    finally
                    {
                        ComUtilities.Release(ref rows);
                        ComUtilities.Release(ref ado);
                        ComUtilities.Release(ref cache);
                        ComUtilities.Release(ref field);
                        ComUtilities.Release(ref fields);
                        ComUtilities.Release(ref cube);
                        ComUtilities.Release(ref cubes);
                        ComUtilities.Release(ref table);
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref worksheet);
                    }
                }
            });
    }

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
        var members = ReadExternalPivotMembers(sheet, pivot, ExternalRelativePeriod);
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
            var snapshot = RequireSuccess(pivots.Read(batch, pivot, sheet));
            Assert.True(snapshot.SnapshotComplete);
            var selection = Assert.Single(snapshot.FieldSelections);
            Assert.True(selection.SelectionStateAvailable);
            Assert.Equal(result.ShowAll, selection.ShowAll);
            Assert.Equal(result.MultipleSelectionEnabled, selection.MultipleSelectionEnabled);
            Assert.Equal(selectedItems.Count == 0 ? null : (int?)selectedItems.Count,
                selection.SelectedMemberCount);
            Assert.Equal(selectedItems.Order(StringComparer.Ordinal),
                selection.SelectedMemberUniqueNames.Order(StringComparer.Ordinal));
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
                Assert.NotNull(sheet);
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

    private Dictionary<string, string> ReadExternalPivotMembers(
        string sheetName, string pivotName, string hierarchy, int level = 1) =>
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
                Assert.NotNull(sheet);
                pivots = sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubes = pivot.CubeFields;
                cube = cubes[hierarchy];
                fields = cube.PivotFields;
                field = fields.Item(level);
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

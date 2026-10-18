using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_ReadReturnsCaptionAndSelectionsFromEveryLevel()
    {
        var (sheet, pivot, _) = CreateConfiguredExternalPivot("ExternalOlapSnapshot");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddRowField(batch, pivot, ExternalCalendar));
        RequireSuccess(pivots.AddValueField(batch, pivot, ExternalAmountMeasure));

        var unfiltered = RequireSuccess(pivots.Read(batch, pivot, sheet));

        Assert.True(unfiltered.IsOlap);
        Assert.True(unfiltered.SnapshotComplete);
        var calendar = Assert.Single(unfiltered.FieldSelections);
        Assert.Equal(ExternalCalendar, calendar.FieldName);
        Assert.Equal("Calendar", calendar.Caption);
        Assert.Equal(PivotFieldArea.Row, calendar.Area);
        Assert.True(calendar.SelectionStateAvailable);
        Assert.True(calendar.ShowAll);
        Assert.Empty(calendar.SelectedMemberUniqueNames);
        Assert.Null(calendar.SelectedMemberCount);

        ExpandExternalCalendar(sheet, pivot);
        string selectedYear = ReadExternalPivotMembers(sheet, pivot, ExternalCalendar)["2025"];
        string selectedMonth = ReadExternalPivotMembers(sheet, pivot, ExternalCalendar, 2)["2025-11"];
        if (string.Equals(Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_SYNTHETIC_FIXTURE"), "true", StringComparison.OrdinalIgnoreCase))
        {
            Assert.Equal("[Sales].[Calendar].[ALL].[AllMember].[2025]", selectedYear);
            Assert.Equal("[Sales].[Calendar].[ALL].[AllMember].[2025].[2025-11]", selectedMonth);
        }
        string[] selectedMembers = new[] { selectedYear, selectedMonth }
            .Order(StringComparer.Ordinal).ToArray();
        SelectExternalYearAndMonth(sheet, pivot, selectedYear, selectedMonth);
        var filtered = RequireSuccess(pivots.Read(batch, pivot, sheet));

        Assert.True(filtered.SnapshotComplete);
        calendar = Assert.Single(filtered.FieldSelections);
        Assert.Equal("Calendar", calendar.Caption);
        Assert.True(calendar.SelectionStateAvailable);
        Assert.False(calendar.ShowAll);
        Assert.Equal(selectedMembers, calendar.SelectedMemberUniqueNames);
        Assert.Equal(2, calendar.SelectedMemberCount);
        Assert.False(calendar.MembersTruncated);
        Assert.True(calendar.FiltersAvailable);
        Assert.Empty(calendar.Filters);

        var bounded = RequireSuccess(pivots.Read(batch, pivot, sheet, maxSelectedMembers: 1));

        Assert.False(bounded.SnapshotComplete);
        calendar = Assert.Single(bounded.FieldSelections);
        Assert.Equal(selectedMembers.Take(1), calendar.SelectedMemberUniqueNames);
        Assert.Equal(2, calendar.SelectedMemberCount);
        Assert.True(calendar.MembersTruncated);
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_ReadReturnsNativeFilterOnSecondLevel()
    {
        var (sheet, pivot, _) = CreateConfiguredExternalPivot("ExternalOlapLevelFilter");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddRowField(batch, pivot, ExternalCalendar));
        RequireSuccess(pivots.AddValueField(batch, pivot, ExternalAmountMeasure));
        AddMonthCaptionFilter(sheet, pivot, "2025-11");

        var snapshot = RequireSuccess(pivots.Read(batch, pivot, sheet));

        Assert.True(snapshot.SnapshotComplete);
        var calendar = Assert.Single(snapshot.FieldSelections);
        Assert.Equal(ExternalCalendar, calendar.FieldName);
        Assert.True(calendar.FiltersAvailable);
        var filter = Assert.Single(calendar.Filters);
        Assert.Equal("[Sales].[Calendar].[Month]", filter.LevelUniqueName);
        Assert.Equal("CaptionEquals", filter.Type);
        Assert.Equal("2025-11", filter.Value1);
        Assert.Null(filter.Value2);
        Assert.Null(filter.DataFieldName);
    }

    private void AddMonthCaptionFilter(string sheetName, string pivotName, string caption)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? fields = null;
            Excel.CubeField? calendar = null;
            Excel.PivotFields? levels = null;
            Excel.PivotField? year = null;
            Excel.PivotField? month = null;
            Excel.PivotFilters? filters = null;
            Excel.PivotFilter? filter = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                pivots = sheet!.PivotTables();
                pivot = pivots.Item(pivotName);
                fields = pivot.CubeFields;
                calendar = fields[ExternalCalendar];
                levels = calendar.PivotFields;
                year = levels.Item(1);
                month = levels.Item(2);
                year.DrilledDown = true;
                filters = month.PivotFilters;
                filter = filters.Add2(Excel.XlPivotFilterType.xlCaptionEquals, Value1: caption);
            }
            finally
            {
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref filters);
                ComUtilities.Release(ref month);
                ComUtilities.Release(ref year);
                ComUtilities.Release(ref levels);
                ComUtilities.Release(ref calendar);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private void SelectExternalYearAndMonth(
        string sheetName, string pivotName, string selectedYear, string selectedMonth)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? fields = null;
            Excel.CubeField? calendar = null;
            Excel.PivotFields? levels = null;
            Excel.PivotField? year = null;
            Excel.PivotField? month = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                pivots = sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                fields = pivot.CubeFields;
                calendar = fields[ExternalCalendar];
                levels = calendar.PivotFields;
                year = levels.Item(1);
                month = levels.Item(2);
                month.VisibleItemsList = new object[] { selectedMonth };
                year.VisibleItemsList = new object[] { selectedYear };
            }
            finally
            {
                ComUtilities.Release(ref month);
                ComUtilities.Release(ref year);
                ComUtilities.Release(ref levels);
                ComUtilities.Release(ref calendar);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

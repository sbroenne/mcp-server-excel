using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [ConfiguredSsasFact]
    [Trait("RunType", "OnDemand")]
    public async Task SsasReportFilter_AdventureWorks_NonstandardAllRoundTripsThroughService()
    {
        const string hierarchy = "[Product].[Category]";
        const string level = "[Product].[Category].[Category]";
        var connection = UniqueConnectionName("SsasAllMember");
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        CreateOlapPivotTable(
            Environment.GetEnvironmentVariable("EXCELMCP_TEST_SSAS_CONNECTION_STRING")!,
            "Adventure Works", connection, sheet);
        var pivot = connection + "_Pivot";
        var batch = _fixture.BatchToken;
        var commands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(commands.AddFilterField(batch, pivot, hierarchy));
        RequireSuccess(commands.AddValueField(batch, pivot, "[Measures].[Internet Sales Amount]"));
        var members = RequireSuccess(_connections.SearchOlapMembers(
            batch, connection, hierarchy, level));
        var bikes = Assert.Single(members.Members, member => member.Caption == "Bikes");
        Assert.Equal("[Product].[Category].&[1]", bikes.UniqueName);

        string allMember = SetNativeSingleSelection(null);
        Assert.False(string.IsNullOrWhiteSpace(allMember));
        Assert.False(allMember.EndsWith(".[All]", StringComparison.OrdinalIgnoreCase),
            "This regression requires a nonstandard All member, not the original name heuristic.");
        double allTotal = AssertSelection(true, null);
        Assert.True(allTotal > 0);

        SetNativeSingleSelection(bikes.UniqueName);
        double bikesTotal = AssertSelection(false, bikes.UniqueName);
        Assert.True(bikesTotal > 0);
        Assert.True(bikesTotal < allTotal);

        var assignedAll = RequireSuccess(commands.SetReportFilter(
            batch, sheet, pivot, hierarchy, [allMember]));
        Assert.Equal([allMember], assignedAll.RequestedItems);
        Assert.True(assignedAll.SelectionStateAvailable);
        Assert.True(assignedAll.ShowAll);
        Assert.False(assignedAll.MultipleSelectionEnabled);
        Assert.Empty(assignedAll.SelectedItems);
        Assert.Equal(allTotal, AssertSelection(true, null));

        var selected = RequireSuccess(commands.SetReportFilter(
            batch, sheet, pivot, hierarchy, [bikes.UniqueName]));
        Assert.False(selected.ShowAll);
        Assert.False(selected.MultipleSelectionEnabled);
        Assert.Equal([bikes.UniqueName], selected.SelectedItems);
        Assert.Equal(bikesTotal, AssertSelection(false, bikes.UniqueName));

        var cleared = RequireSuccess(commands.SetReportFilter(batch, sheet, pivot, hierarchy, []));
        Assert.True(cleared.ShowAll);
        Assert.True(cleared.SelectionStateAvailable);
        Assert.False(cleared.MultipleSelectionEnabled);
        Assert.Empty(cleared.SelectedItems);
        Assert.Equal(allTotal, AssertSelection(true, null));

        await _fixture.SaveAndReopenAsync();
        RequireSuccess(commands.Refresh(batch, pivot));
        Assert.Equal(allTotal, AssertSelection(true, null));
        RequireSuccess(commands.SetReportFilter(batch, sheet, pivot, hierarchy, [bikes.UniqueName]));
        Assert.Equal(bikesTotal, AssertSelection(false, bikes.UniqueName));
        RequireSuccess(commands.SetReportFilter(batch, sheet, pivot, hierarchy, []));
        Assert.Equal(allTotal, AssertSelection(true, null));

        double AssertSelection(bool showAll, string? member)
        {
            var filter = RequireSuccess(commands.GetReportFilter(batch, sheet, pivot, hierarchy));
            Assert.True(filter.SelectionStateAvailable);
            Assert.Equal(PivotFieldArea.Filter, filter.FieldArea);
            Assert.Equal(showAll, filter.ShowAll);
            Assert.False(filter.MultipleSelectionEnabled);
            Assert.Equal(member is null ? [] : new[] { member }, filter.SelectedItems);
            var snapshot = RequireSuccess(commands.Read(batch, pivot, sheet));
            Assert.True(snapshot.SnapshotComplete);
            var field = Assert.Single(snapshot.FieldSelections);
            Assert.Equal(hierarchy, field.FieldName);
            Assert.Equal(PivotFieldArea.Filter, field.Area);
            Assert.True(field.SelectionStateAvailable);
            Assert.Equal(showAll, field.ShowAll);
            Assert.False(field.MultipleSelectionEnabled);
            Assert.Equal(showAll ? null : (int?)1, field.SelectedMemberCount);
            Assert.Equal(member is null ? [] : new[] { member }, field.SelectedMemberUniqueNames);
            Assert.False(field.MembersTruncated);
            var data = RequireSuccess(commands.GetData(batch, pivot));
            return Convert.ToDouble(
                Assert.Single(data.Values.SelectMany(row => row), value => value is double or int),
                CultureInfo.InvariantCulture);
        }

        string SetNativeSingleSelection(string? selectedMember) =>
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
                    cube = cubes[hierarchy];
                    fields = cube.PivotFields;
                    field = fields.Item(1);
                    cache = table.PivotCache();
                    ado = (ADODB._Connection)cache.ADOConnection;
                    rows = ado.OpenSchema((ADODB.SchemaEnum)34, new object?[]
                        { null, null, "Adventure Works", null, null, hierarchy }, Type.Missing);
                    Assert.False(Convert.ToBoolean(rows.EOF, CultureInfo.InvariantCulture));
                    columns = rows.Fields;
                    all = columns["ALL_MEMBER"];
                    string nativeAll = Convert.ToString((object)all.Value, CultureInfo.InvariantCulture)!;
                    Assert.False(string.IsNullOrWhiteSpace(nativeAll));
                    cube.EnableMultiplePageItems = false;
                    field.CurrentPageName = selectedMember ?? nativeAll;
                    Assert.Equal(selectedMember ?? nativeAll, field.CurrentPageName);
                    Assert.False(cube.EnableMultiplePageItems);
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
}

[AttributeUsage(AttributeTargets.Method)]
public sealed class ConfiguredSsasFactAttribute : FactAttribute
{
    public ConfiguredSsasFactAttribute()
    {
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("EXCELMCP_TEST_SSAS_CONNECTION_STRING")))
            Skip = "Configure the SSAS AdventureWorksDW sample connection to run this regression.";
    }
}

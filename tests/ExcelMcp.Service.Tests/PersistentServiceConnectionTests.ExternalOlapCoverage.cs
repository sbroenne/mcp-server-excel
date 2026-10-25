using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    private const string ExternalCalendar = "[Sales].[Calendar]";
    private const string ExternalRelativePeriod = "[Sales].[RelativePeriod]";

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public async Task ExternalOlapPivotTable_SaveReopenPreservesConnectionLayoutSelectionAndChart()
    {
        var (sheet, pivot, connection) = CreateConfiguredExternalPivot("ExternalOlapPersist");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddRowField(batch, pivot, ExternalCalendar));
        RequireSuccess(pivots.AddValueField(batch, pivot, ExternalAmountMeasure, customName: "Saved Amount"));
        RequireSuccess(pivots.SetLayoutOptions(batch, pivot,
            new PivotLayoutOptions { RowLayout = 1, StyleName = "PivotStyleMedium9", PreserveFormatting = true }));
        RequireSuccess(pivots.SetCacheOptions(batch, pivot, refreshOnFileOpen: false));
        var slicers = _fixture.CreateCommands<ISlicerCommands>();
        var slicerName = connection + "_Slicer";
        RequireSuccess(slicers.CreateSlicer(batch, pivot, ExternalCalendar, slicerName, sheet, "H3"));
        RequireSuccess(slicers.SetSlicerSelection(batch, slicerName, ["2025"]));
        var charts = _fixture.CreateCommands<IChartCommands>();
        var chart = RequireSuccess(charts.CreateFromPivotTable(batch, pivot, sheet,
            ChartType.ColumnClustered, chartName: connection + "_Chart"));
        var before = ReadExternalPivotState(sheet, pivot);
        AssertExternalGrandTotal(pivot, 222);
        AssertChart();

        await _fixture.SaveAndReopenAsync();

        Assert.Equal(before, ReadExternalPivotState(sheet, pivot));
        Assert.Equal(connection, RequireSuccess(pivots.GetConnection(batch, sheet, pivot)).ConnectionName);
        Assert.False(RequireSuccess(pivots.GetCacheOptions(batch, pivot)).RefreshOnFileOpen);
        var layout = RequireSuccess(pivots.GetLayoutOptions(batch, pivot));
        Assert.Equal("PivotStyleMedium9", layout.StyleName);
        Assert.True(layout.PreserveFormatting);
        Assert.All(layout.RowFields, field => Assert.Equal(1, field.RowLayout));
        var value = Assert.Single(RequireSuccess(pivots.ListFields(batch, pivot)).ValueFields);
        Assert.Equal(ExternalAmountMeasure, value.FieldName);
        Assert.Equal("Saved Amount", RequireSuccess(pivots.GetData(batch, pivot)).Values[0][1]);
        AssertChart();
        RequireSuccess(pivots.Refresh(batch, pivot));
        AssertExternalGrandTotal(pivot, 222);
        RequireSuccess(slicers.SetSlicerSelection(batch, slicerName, []));
        AssertExternalGrandTotal(pivot, 300);
        RequireSuccess(charts.Delete(batch, chart.ChartName));
        RequireSuccess(slicers.DeleteSlicer(batch, slicerName));

        void AssertChart()
        {
            var read = RequireSuccess(charts.Read(batch, chart.ChartName));
            Assert.True(read.IsPivotChart);
            Assert.Equal(pivot, read.LinkedPivotTable);
            var series = Assert.Single(read.Series);
            Assert.Equal(["2025"], series.Categories.Select(item => item?.ToString()));
            Assert.Equal([222d], series.Values.Select(item => Convert.ToDouble(item, CultureInfo.InvariantCulture)));
        }
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_MultiLevelRowsColumnsAndFilterReturnExactMonthlyAmounts()
    {
        var (sheet, pivot, connection) = CreateConfiguredExternalPivot("ExternalOlapLevels");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddRowField(batch, pivot, ExternalCalendar));
        RequireSuccess(pivots.AddColumnField(batch, pivot, ExternalRelativePeriod));
        RequireSuccess(pivots.AddValueField(batch, pivot, ExternalAmountMeasure));
        ExpandExternalCalendar(sheet, pivot);
        RequireSuccess(pivots.SetSubtotals(batch, pivot, ExternalCalendar, false));
        RequireSuccess(pivots.SetLayoutOptions(batch, pivot,
            new PivotLayoutOptions { RowLayout = 1, RepeatLabels = true }));
        var layout = RequireSuccess(pivots.GetLayoutOptions(batch, pivot));
        Assert.Equal(2, layout.RowFields.Count);
        Assert.All(layout.RowFields, field =>
        {
            Assert.Equal(1, field.RowLayout);
            Assert.True(field.RepeatLabels);
        });
        var fields = RequireSuccess(pivots.ListFields(batch, pivot));
        Assert.Contains(fields.Fields, field => field.Name == ExternalCalendar && field.Area == PivotFieldArea.Row);
        Assert.Contains(fields.Fields, field => field.Name == ExternalRelativePeriod && field.Area == PivotFieldArea.Column);
        var values = RequireSuccess(pivots.GetData(batch, pivot)).Values;
        Assert.Equal(24, values.Count(row => row.Count > 1 && row[1]?.ToString()?.StartsWith("202", StringComparison.Ordinal) == true));
        for (int index = 0; index < 24; index++)
        {
            var month = $"{2024 + index / 12}-{index % 12 + 1:00}";
            var row = Assert.Single(values, row => row.Count > 1 && row[1]?.ToString() == month);
            Assert.Equal(month[..4], row[0]?.ToString());
            Assert.Equal(index + 1d, Convert.ToDouble(row[^1], CultureInfo.InvariantCulture));
        }
        AssertExternalGrandTotal(pivot, 300);
        RequireSuccess(pivots.RemoveField(batch, pivot, ExternalRelativePeriod));
        var filter = RequireSuccess(pivots.AddFilterField(batch, pivot, ExternalRelativePeriod));
        Assert.Equal(PivotFieldArea.Filter, filter.Area);
        var slicers = _fixture.CreateCommands<ISlicerCommands>();
        var slicerName = connection + "_Period";
        var slicer = RequireSuccess(slicers.CreateSlicer(batch, pivot, ExternalRelativePeriod, slicerName, sheet, "K3"));
        Assert.Equal(["Current Month", "Earlier Months", "Next Month", "Prior Month"],
            slicer.AvailableItems.Order(StringComparer.Ordinal));
        RequireSuccess(slicers.SetSlicerSelection(batch, slicerName, ["Current Month"]));
        var filtered = RequireSuccess(pivots.GetData(batch, pivot)).Values;
        Assert.Contains(filtered, row => row.Any(cell => cell?.ToString() == "2025-11"));
        Assert.DoesNotContain(filtered, row => row.Any(cell => cell?.ToString() == "2025-10"));
        AssertExternalGrandTotal(pivot, 23);
        RequireSuccess(slicers.SetSlicerSelection(batch, slicerName, []));
        AssertExternalGrandTotal(pivot, 300);
        RequireSuccess(slicers.DeleteSlicer(batch, slicerName));
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public async Task ExternalOlapPivotTable_ConnectionGuardsPreserveSelectedAndOtherPivot()
    {
        var (sheet, pivot, connection) = CreateConfiguredExternalPivot("ExternalOlapGuards");
        var (otherSheet, otherPivot, target) = CreateConfiguredExternalPivot("ExternalOlapOther");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        foreach (var name in new[] { pivot, otherPivot })
        {
            RequireSuccess(pivots.AddRowField(batch, name, ExternalCalendar));
            RequireSuccess(pivots.AddValueField(batch, name, ExternalAmountMeasure));
        }
        var before = ReadExternalPivotState(sheet, pivot);
        var otherBefore = ReadExternalPivotState(otherSheet, otherPivot);
        var noOp = RequireSuccess(pivots.SetConnection(batch, sheet, pivot, connection));
        Assert.Equal(connection, noOp.ConnectionName);
        Assert.Equal(before, ReadExternalPivotState(sheet, pivot));
        var missing = await _fixture.SendForFailureAsync("pivottable.set-connection",
            new { sheetName = sheet, pivotTableName = pivot, connectionName = "MissingExternalConnection" });
        Assert.Contains("not found", missing.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, ReadExternalPivotState(sheet, pivot));
        var slicers = _fixture.CreateCommands<ISlicerCommands>();
        var slicerName = connection + "_Guard";
        RequireSuccess(slicers.CreateSlicer(batch, pivot, ExternalCalendar, slicerName, sheet, "H3"));
        var connectedBefore = ReadExternalPivotState(sheet, pivot);
        var rejected = await _fixture.SendForFailureAsync("pivottable.set-connection",
            new { sheetName = sheet, pivotTableName = pivot, connectionName = target });
        Assert.Contains("slicer", rejected.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(connectedBefore, ReadExternalPivotState(sheet, pivot));
        Assert.Equal(otherBefore, ReadExternalPivotState(otherSheet, otherPivot));
        AssertExternalGrandTotal(pivot, 300);
        AssertExternalGrandTotal(otherPivot, 300);
        RequireSuccess(slicers.DeleteSlicer(batch, slicerName));
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_ChangeConnectionReportsActualProviderResult()
    {
        var (sheet, pivot, original) = CreateConfiguredExternalPivot("ExternalOlapChange");
        var (otherSheet, otherPivot, target) = CreateConfiguredExternalPivot("ExternalOlapTarget");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        foreach (var name in new[] { pivot, otherPivot })
        {
            RequireSuccess(pivots.AddRowField(batch, name, ExternalCalendar));
            RequireSuccess(pivots.AddValueField(batch, name, ExternalAmountMeasure));
        }
        var otherBefore = ReadExternalPivotState(otherSheet, otherPivot);

        var changed = RequireSuccess(pivots.SetConnection(batch, sheet, pivot, target));

        Assert.Equal(target, changed.ConnectionName);
        Assert.NotEqual(original, changed.ConnectionName);
        RequireSuccess(pivots.Refresh(batch, pivot));
        AssertExternalGrandTotal(pivot, 300);
        Assert.Equal(otherBefore, ReadExternalPivotState(otherSheet, otherPivot));
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_InvalidCubeRefreshReportsFailureAndAllowsExplicitRepair()
    {
        var (sheet, pivot, connection) = CreateConfiguredExternalPivot("ExternalOlapRefreshFailure");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddRowField(batch, pivot, ExternalCalendar));
        RequireSuccess(pivots.AddValueField(batch, pivot, ExternalAmountMeasure));
        var before = ReadExternalPivotState(sheet, pivot);
        var cube = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_CUBE")!;
        PersistentServiceCleanupFailures.Run(() =>
        {
            SetCubeCommand("MissingExcelMcpCube");
            var error = Assert.Throws<InvalidOperationException>(
                () => pivots.Refresh(batch, pivot, TimeSpan.FromSeconds(15)));
            Assert.Contains("pivottable.refresh failed [ComInterop/COMException]", error.Message);
            Assert.Equal(before, ReadExternalPivotState(sheet, pivot));
            SetCubeCommand(cube);
            RequireSuccess(_connections.Refresh(batch, connection));
            RequireSuccess(pivots.Refresh(batch, pivot));
            AssertExternalGrandTotal(pivot, 300);
        }, () => SetCubeCommand(cube));

        void SetCubeCommand(string command)
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? worksheet = null;
                Excel.PivotTables? tables = null;
                Excel.PivotTable? table = null;
                Excel.PivotCache? cache = null;
                try
                {
                    worksheet = ComUtilities.FindSheet(context.Book, sheet);
                    Assert.NotNull(worksheet);
                    tables = worksheet.PivotTables();
                    table = tables.Item(pivot);
                    cache = table.PivotCache();
                    cache.CommandText = command;
                    Assert.Equal(command, ((object)cache.CommandText));
                }
                finally
                {
                    ComUtilities.Release(ref cache);
                    ComUtilities.Release(ref table);
                    ComUtilities.Release(ref tables);
                    ComUtilities.Release(ref worksheet);
                }
            });
        }
    }

    private string ReadExternalPivotState(string sheet, string pivot)
    {
        var commands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var connection = RequireSuccess(commands.GetConnection(_fixture.BatchToken, sheet, pivot));
        var fields = RequireSuccess(commands.ListFields(_fixture.BatchToken, pivot));
        var data = RequireSuccess(commands.GetData(_fixture.BatchToken, pivot));
        return JsonSerializer.Serialize(new
        {
            connection.ConnectionName,
            connection.IsExternal,
            connection.IsOlap,
            connection.SharedPivotTables,
            connection.ConnectedSlicerCaches,
            fields.Fields,
            fields.ValueFields,
            data.Values
        });
    }

    private void AssertExternalGrandTotal(string pivot, double expected)
    {
        var commands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var data = RequireSuccess(commands.GetData(_fixture.BatchToken, pivot));
        Assert.Equal(expected, Convert.ToDouble(data.Values[^1][^1], CultureInfo.InvariantCulture));
    }

    private void ExpandExternalCalendar(string sheetName, string pivotName)
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
                year.DrilledDown = true;
                Assert.True(year.DrilledDown);
            }
            finally
            {
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

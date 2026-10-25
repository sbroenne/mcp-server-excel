using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using System.Text.Json;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    private const string ExternalAmountMeasure = "[Measures].[Amount.SUM]";

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_FieldCalcSlicerAndChartCommandsUseServerCube()
    {
        var (sheetName, pivotName, connectionName) = CreateConfiguredExternalPivot("ExternalOlapPivot");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        string hierarchy = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_HIERARCHY")!;
        var fields = RequireSuccess(pivots.ListFields(batch, pivotName));
        Assert.Contains(fields.Fields, field => field.Name == hierarchy);
        Assert.Contains(fields.Fields, field => field.Name == ExternalAmountMeasure);

        var connection = RequireSuccess(pivots.GetConnection(batch, sheetName, pivotName));
        Assert.True(connection.IsOlap);
        Assert.True(connection.IsExternal);
        Assert.False(connection.IsDataModel);
        Assert.Equal(connectionName, connection.ConnectionName);
        Assert.True(RequireSuccess(pivots.GetCacheOptions(batch, pivotName)).IsOlap);

        var row = RequireSuccess(pivots.AddRowField(batch, pivotName, hierarchy));
        Assert.Equal(PivotFieldArea.Row, row.Area);
        var value = RequireSuccess(pivots.AddValueField(batch, pivotName, ExternalAmountMeasure));
        Assert.Equal("Amount.SUM", value.FieldName);
        Assert.Equal(PivotFieldArea.Value, value.Area);
        RequireSuccess(pivots.Refresh(batch, pivotName));

        var placed = RequireSuccess(pivots.ListFields(batch, pivotName));
        Assert.Contains(placed.Fields, field => field.Name == hierarchy && field.Area == PivotFieldArea.Row);
        Assert.Contains(placed.Fields, field => field.Name == ExternalAmountMeasure && field.Area == PivotFieldArea.Value);
        var data = RequireSuccess(pivots.GetData(batch, pivotName));
        Assert.Contains(data.Values, cells => cells[0]?.ToString() == "2024"
            && Convert.ToDouble(cells[1], System.Globalization.CultureInfo.InvariantCulture) == 78d);
        Assert.Contains(data.Values, cells => cells[0]?.ToString() == "2025"
            && Convert.ToDouble(cells[1], System.Globalization.CultureInfo.InvariantCulture) == 222d);
        Assert.Equal(300d, Convert.ToDouble(data.Values[^1][1], System.Globalization.CultureInfo.InvariantCulture));

        RequireSuccess(pivots.SortField(batch, pivotName, hierarchy, SortDirection.Descending));
        var sorted = RequireSuccess(pivots.GetData(batch, pivotName));
        Assert.Equal(["2025", "2024"], sorted.Values.Select(cells => cells[0]?.ToString())
            .Where(label => label is "2024" or "2025"));
        var formatted = RequireSuccess(pivots.SetFieldFormat(batch, pivotName, ExternalAmountMeasure, "0%"));
        Assert.Equal("0%", formatted.NumberFormat);

        const string calculatedName = "[Measures].[ExcelMcpDouble]";
        RequireSuccess(pivots.CreateCalculatedMember(batch, pivotName, calculatedName, ExternalAmountMeasure + " * 2"));
        Assert.Contains(RequireSuccess(pivots.ListCalculatedMembers(batch, pivotName)).CalculatedMembers,
            member => member.Name == calculatedName && member.Formula.Contains(ExternalAmountMeasure, StringComparison.Ordinal));
        RequireSuccess(pivots.Refresh(batch, pivotName));
        RequireSuccess(pivots.AddValueField(batch, pivotName, calculatedName));
        var calculated = RequireSuccess(pivots.GetData(batch, pivotName));
        Assert.Equal(600d, Convert.ToDouble(calculated.Values[^1][2], System.Globalization.CultureInfo.InvariantCulture));
        RequireSuccess(pivots.RemoveField(batch, pivotName, calculatedName));
        RequireSuccess(pivots.DeleteCalculatedMember(batch, pivotName, calculatedName));
        Assert.DoesNotContain(RequireSuccess(pivots.ListCalculatedMembers(batch, pivotName)).CalculatedMembers,
            member => member.Name == calculatedName);

        var slicers = _fixture.CreateCommands<ISlicerCommands>();
        var slicerName = connectionName + "_Slicer";
        var slicer = RequireSuccess(slicers.CreateSlicer(batch, pivotName, hierarchy, slicerName, sheetName, "H3"));
        Assert.Equal([pivotName], slicer.ConnectedPivotTables);
        Assert.Contains("2024", slicer.AvailableItems);
        Assert.Contains("2025", slicer.AvailableItems);
        RequireSuccess(slicers.SetSlicerSelection(batch, slicerName, ["2024"]));
        Assert.Equal(78d, Convert.ToDouble(RequireSuccess(pivots.GetData(batch, pivotName)).Values[^1][1],
            System.Globalization.CultureInfo.InvariantCulture));
        var invalidSelection = Assert.Throws<ArgumentException>(
            () => slicers.SetSlicerSelection(batch, slicerName, ["NoSuchCalendarMember"]));
        Assert.Contains("not found", invalidSelection.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(78d, Convert.ToDouble(RequireSuccess(pivots.GetData(batch, pivotName)).Values[^1][1],
            System.Globalization.CultureInfo.InvariantCulture));
        RequireSuccess(slicers.SetSlicerSelection(batch, slicerName, []));
        Assert.Equal(300d, Convert.ToDouble(RequireSuccess(pivots.GetData(batch, pivotName)).Values[^1][1],
            System.Globalization.CultureInfo.InvariantCulture));
        RequireSuccess(slicers.DeleteSlicer(batch, slicerName));
        Assert.DoesNotContain(RequireSuccess(slicers.ListSlicers(batch)).Slicers, item => item.Name == slicerName);

        var charts = _fixture.CreateCommands<IChartCommands>();
        var chart = RequireSuccess(charts.CreateFromPivotTable(
            batch, pivotName, sheetName, ChartType.ColumnClustered, chartName: connectionName + "_Chart"));
        var chartRead = RequireSuccess(charts.Read(batch, chart.ChartName));
        Assert.Equal(chart.ChartName, chartRead.Name);
        Assert.Equal(ChartType.ColumnClustered, chartRead.ChartType);
        Assert.True(chartRead.IsPivotChart);
        Assert.Equal(pivotName, chartRead.LinkedPivotTable);
        var series = Assert.Single(chartRead.Series);
        Assert.Equal(["2025", "2024"], series.Categories.Select(item => item?.ToString()));
        Assert.Equal([222d, 78d], series.Values.Select(item =>
            Convert.ToDouble(item, System.Globalization.CultureInfo.InvariantCulture)));
        RequireSuccess(charts.Delete(batch, chart.ChartName));
        var removed = RequireSuccess(pivots.RemoveField(batch, pivotName, hierarchy));
        Assert.Equal(PivotFieldArea.Hidden, removed.Area);
        Assert.Contains(RequireSuccess(pivots.ListFields(batch, pivotName)).Fields,
            field => field.Name == hierarchy && field.Area == PivotFieldArea.Hidden);
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_ChangeMeasureAndGroupingCommandsReportCubeLimits()
    {
        var (_, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapLimits");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        string hierarchy = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_HIERARCHY")!;
        RequireSuccess(pivots.AddRowField(batch, pivotName, hierarchy));
        RequireSuccess(pivots.AddValueField(batch, pivotName, ExternalAmountMeasure));
        RequireSuccess(pivots.Refresh(batch, pivotName));
        var before = RequireSuccess(pivots.GetData(batch, pivotName)).Values;
        var attempts = new (Func<PivotFieldResult> Run, string ExpectedText)[]
        {
            (() => pivots.SetFieldFunction(batch, pivotName, ExternalAmountMeasure, AggregationFunction.Average), "cube server"),
            (() => pivots.SetFieldName(batch, pivotName, ExternalAmountMeasure, "Renamed"), "cube"),
            (() => pivots.GroupByDate(batch, pivotName, hierarchy, DateGroupingInterval.Months), "cube server"),
            (() => pivots.GroupByNumeric(batch, pivotName, ExternalAmountMeasure, 0, 100, 10), "cube server")
        };
        foreach (var (run, expectedText) in attempts)
        {
            var result = run();
            Assert.False(result.Success);
            Assert.Contains(expectedText, result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("workbook has no Data Model", result.ErrorMessage ?? "");
            var retained = RequireSuccess(pivots.GetData(batch, pivotName)).Values;
            Assert.Equal(before.Count, retained.Count);
            for (int index = 0; index < before.Count; index++) { Assert.Equal(before[index], retained[index]); }
        }
        Assert.Contains(RequireSuccess(pivots.ListFields(batch, pivotName)).Fields,
            field => field.Name == ExternalAmountMeasure && field.Area == PivotFieldArea.Value);
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_BareDottedMeasureIsExactAndIdempotent()
    {
        var (_, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapMeasure");
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var added = RequireSuccess(pivots.AddValueField(_fixture.BatchToken, pivotName, "Amount.SUM"));
        Assert.Equal("Amount.SUM", added.FieldName);
        Assert.Equal(PivotFieldArea.Value, added.Area);
        var repeated = RequireSuccess(pivots.AddValueField(_fixture.BatchToken, pivotName, ExternalAmountMeasure));
        Assert.Equal("Amount.SUM", repeated.FieldName);
        var fields = RequireSuccess(pivots.ListFields(_fixture.BatchToken, pivotName));
        Assert.Single(fields.Fields, field => field.Area == PivotFieldArea.Value && field.Name == ExternalAmountMeasure);
        Assert.Single(fields.ValueFields);
        Assert.Equal(300d, Convert.ToDouble(RequireSuccess(pivots.GetData(_fixture.BatchToken, pivotName)).Values[^1][0],
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_FieldReadsReturnActualCaptionAndPosition()
    {
        var (_, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapFields");
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        string hierarchy = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_HIERARCHY")!;
        var row = RequireSuccess(pivots.AddRowField(_fixture.BatchToken, pivotName, hierarchy));
        var listed = Assert.Single(RequireSuccess(pivots.ListFields(_fixture.BatchToken, pivotName)).Fields,
            field => field.Name == hierarchy);
        var read = Assert.Single(RequireSuccess(pivots.Read(_fixture.BatchToken, pivotName)).Fields,
            field => field.Name == hierarchy);
        Assert.NotEmpty(row.CustomName);
        Assert.Equal(1, row.Position);
        Assert.Equal(row.CustomName, listed.CustomName);
        Assert.Equal(row.Position, listed.Position);
        Assert.Equal(row.CustomName, read.CustomName);
        Assert.Equal(row.Position, read.Position);
        Assert.Equal(PivotFieldArea.Row, read.Area);
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_SubtotalsAcceptHierarchyAndPreserveValues()
    {
        var (sheet, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapSubtotals");
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        string hierarchy = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_HIERARCHY")!;
        RequireSuccess(pivots.AddRowField(_fixture.BatchToken, pivotName, hierarchy));
        RequireSuccess(pivots.AddValueField(_fixture.BatchToken, pivotName, ExternalAmountMeasure));
        RequireSuccess(pivots.SetSubtotals(_fixture.BatchToken, pivotName, hierarchy, false));
        AssertExternalSubtotals(sheet, pivotName, hierarchy, false);
        RequireSuccess(pivots.SetSubtotals(_fixture.BatchToken, pivotName, hierarchy, true));
        AssertExternalSubtotals(sheet, pivotName, hierarchy, true);
        Assert.Equal(300d, Convert.ToDouble(RequireSuccess(pivots.GetData(_fixture.BatchToken, pivotName)).Values[^1][1],
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_DisabledRefreshCannotReportSuccess()
    {
        var (_, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapRefresh");
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddValueField(_fixture.BatchToken, pivotName, ExternalAmountMeasure));
        var before = RequireSuccess(pivots.GetData(_fixture.BatchToken, pivotName)).Values;
        Assert.False(RequireSuccess(pivots.SetCacheOptions(
            _fixture.BatchToken, pivotName, enableRefresh: false)).EnableRefresh);

        var error = Assert.Throws<InvalidOperationException>(() => pivots.Refresh(_fixture.BatchToken, pivotName));
        Assert.Contains("refresh", error.Message, StringComparison.OrdinalIgnoreCase);
        var retained = RequireSuccess(pivots.GetData(_fixture.BatchToken, pivotName)).Values;
        Assert.Equal(before.Count, retained.Count);
        for (int index = 0; index < before.Count; index++) { Assert.Equal(before[index], retained[index]); }
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_CannotModifySameNamedWorkbookMeasure()
    {
        var batch = _fixture.BatchToken;
        var modelSheet = _fixture.CreateTestSheet(batch);
        var tableName = "LocalSales" + Guid.NewGuid().ToString("N")[..8];
        RequireSuccess(_commands.SetValues(batch, modelSheet, "A1:B3",
            [["Category", "Amount"], ["One", 10], ["Two", 20]]));
        var tables = _fixture.CreateCommands<ITableCommands>();
        RequireSuccess(tables.Create(batch, modelSheet, tableName, "A1:B3"));
        _fixture.RegisterTableForCleanup(tableName);
        RequireSuccess(tables.AddToDataModel(batch, tableName));
        _fixture.RegisterDataModelTableForCleanup(tableName);
        var model = _fixture.CreateCommands<IDataModelCommands>();
        string formula = $"SUM('{tableName}'[Amount])";
        RequireSuccess(model.CreateMeasure(batch, tableName, "Amount.SUM", formula));
        _fixture.RegisterDataModelMeasureForCleanup("Amount.SUM");
        var (_, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapCollision");
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivots.AddValueField(batch, pivotName, ExternalAmountMeasure));

        var result = pivots.SetFieldFunction(batch, pivotName, ExternalAmountMeasure, AggregationFunction.Count);

        Assert.False(result.Success);
        Assert.Contains("cube server", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var localMeasure = Assert.Single(RequireSuccess(model.ListMeasures(batch, tableName)).Measures,
            measure => measure.Name == "Amount.SUM");
        Assert.Equal(formula, localMeasure.FormulaPreview);
        Assert.Equal(300d, Convert.ToDouble(RequireSuccess(pivots.GetData(batch, pivotName)).Values[^1][0],
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_ExistingMeasureCustomNameMatchesDisplayedCaption()
    {
        var (_, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapCaption");
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var added = RequireSuccess(pivots.AddValueField(
            _fixture.BatchToken, pivotName, ExternalAmountMeasure, customName: "Requested Amount"));
        var data = RequireSuccess(pivots.GetData(_fixture.BatchToken, pivotName));
        Assert.Equal("Requested Amount", added.CustomName);
        Assert.Equal(added.CustomName, data.Values[0][0]);
        Assert.Equal(300d, Convert.ToDouble(data.Values[^1][0], System.Globalization.CultureInfo.InvariantCulture));
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public async Task ExternalOlapPivotTable_UnsupportedDepthActionsPreserveFieldsDataAndSheets()
    {
        var (_, pivotName, _) = CreateConfiguredExternalPivot("ExternalOlapDepth");
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        string hierarchy = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_HIERARCHY")!;
        RequireSuccess(pivots.AddRowField(batch, pivotName, hierarchy));
        RequireSuccess(pivots.AddValueField(batch, pivotName, ExternalAmountMeasure));
        var before = ReadState();
        var attempts = new (string Action, object Arguments, string ExpectedText)[]
        {
            ("pivottablefield.group-items", new { pivotTableName = pivotName, fieldName = hierarchy,
                itemNames = new[] { "2024", "2025" }, groupName = "Years" }, "Manual item grouping"),
            ("pivottablefield.ungroup-field", new { pivotTableName = pivotName,
                groupedFieldName = hierarchy }, "Manual ungrouping"),
            ("pivottable.drill-through", new { pivotTableName = pivotName, cellAddress = "B5" }, "provider-dependent"),
            ("pivottablefield.get-item-expansion", new { pivotTableName = pivotName,
                fieldName = hierarchy, itemName = "2024" }, "OLAP"),
            ("pivottablefield.set-item-expansion", new { pivotTableName = pivotName,
                fieldName = hierarchy, itemName = "2024", expanded = false }, "OLAP"),
            ("pivottable.get-source", new { pivotTableName = pivotName }, "OLAP")
        };
        foreach (var (action, arguments, expectedText) in attempts)
        {
            var error = await _fixture.SendForFailureAsync(action, arguments);
            Assert.False(error.Success);
            Assert.Contains(expectedText, error.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(before, ReadState());
        }

        string ReadState()
        {
            int sheetCount = 0;
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Sheets? sheets = null;
                try { sheets = context.Book.Worksheets; sheetCount = sheets.Count; }
                finally { ComUtilities.Release(ref sheets); }
            });
            var fields = RequireSuccess(pivots.ListFields(batch, pivotName));
            var data = RequireSuccess(pivots.GetData(batch, pivotName));
            Assert.Equal(300d, Convert.ToDouble(data.Values[^1][1], System.Globalization.CultureInfo.InvariantCulture));
            return JsonSerializer.Serialize(new { sheetCount, fields.Fields, fields.ValueFields, data.Values });
        }
    }

    private (string Sheet, string Pivot, string Connection) CreateConfiguredExternalPivot(string prefix)
    {
        var connection = UniqueConnectionName(prefix);
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        CreateOlapPivotTable(
            Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_CONNECTION_STRING")!,
            Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_CUBE")!, connection, sheet);
        return (sheet, connection + "_Pivot", connection);
    }

    private void AssertExternalSubtotals(string sheetName, string pivotName, string hierarchy, bool expected)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubes = null;
            Excel.CubeField? cube = null;
            Excel.PivotFields? fields = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                pivots = sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cubes = pivot.CubeFields;
                cube = cubes[hierarchy];
                fields = cube.PivotFields;
                Assert.True(fields.Count > 0);
                for (int index = 1; index <= fields.Count; index++)
                {
                    Excel.PivotField? field = null;
                    try
                    {
                        field = fields.Item(index);
                        Assert.Equal(expected, Convert.ToBoolean(((object)field.Subtotals[1]), System.Globalization.CultureInfo.InvariantCulture));
                    }
                    finally { ComUtilities.Release(ref field); }
                }
            }
            finally
            {
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref cube);
                ComUtilities.Release(ref cubes);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}

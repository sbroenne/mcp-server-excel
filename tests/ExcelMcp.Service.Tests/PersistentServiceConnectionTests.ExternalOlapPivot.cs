using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    // Data Model tests cover OLAP PivotTables backed by the workbook itself; this one
    // proves the same commands work when the PivotTable reads from a real cube server.
    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapPivotTable_FieldCalcSlicerAndChartCommandsUseServerCube()
    {
        string connectionString = Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_CONNECTION_STRING")!;
        string cubeName = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_CUBE")!;
        string hierarchy = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_HIERARCHY")!;
        var connectionName = UniqueConnectionName("ExternalOlapPivot");
        var pivotName = connectionName + "_Pivot";
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var batch = _fixture.BatchToken;
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();

        try
        {
            CreateOlapPivotTable(connectionString, cubeName, connectionName, sheetName);

            var fields = RequireSuccess(pivots.ListFields(batch, pivotName));
            Assert.Contains(fields.Fields, field => field.Name == hierarchy);
            var measure = fields.Fields
                .Select(field => field.Name)
                .First(name => name.StartsWith("[Measures].", StringComparison.OrdinalIgnoreCase));

            var connection = RequireSuccess(pivots.GetConnection(batch, sheetName, pivotName));
            Assert.True(connection.IsOlap);
            Assert.Equal(connectionName, connection.ConnectionName);
            Assert.True(RequireSuccess(pivots.GetCacheOptions(batch, pivotName)).IsOlap);

            var row = RequireSuccess(pivots.AddRowField(batch, pivotName, hierarchy));
            Assert.Equal(PivotFieldArea.Row, row.Area);
            var value = RequireSuccess(pivots.AddValueField(batch, pivotName, measure));
            Assert.Equal(PivotFieldArea.Value, value.Area);
            RequireSuccess(pivots.Refresh(batch, pivotName));

            var placed = RequireSuccess(pivots.ListFields(batch, pivotName));
            Assert.Contains(placed.Fields, field => field.Name == hierarchy && field.Area == PivotFieldArea.Row);
            Assert.Contains(placed.Fields, field => field.Name == measure && field.Area == PivotFieldArea.Value);
            var data = RequireSuccess(pivots.GetData(batch, pivotName));
            Assert.True(data.Values.Count >= 2, "The server cube should return at least one member row.");
            Assert.Contains(data.Values.Skip(1), cells => cells.Skip(1).Any(cell => cell is double));

            RequireSuccess(pivots.SortField(batch, pivotName, hierarchy, SortDirection.Descending));
            RequireSuccess(pivots.SetFieldFormat(batch, pivotName, measure, "#,##0.00"));

            var calculatedName = "[Measures].[ExcelMcpDouble]";
            RequireSuccess(pivots.CreateCalculatedMember(batch, pivotName, calculatedName, measure + " * 2"));
            Assert.Contains(RequireSuccess(pivots.ListCalculatedMembers(batch, pivotName)).CalculatedMembers,
                member => member.Name == calculatedName);
            RequireSuccess(pivots.DeleteCalculatedMember(batch, pivotName, calculatedName));
            Assert.DoesNotContain(RequireSuccess(pivots.ListCalculatedMembers(batch, pivotName)).CalculatedMembers,
                member => member.Name == calculatedName);

            var slicers = _fixture.CreateCommands<ISlicerCommands>();
            var slicerName = connectionName + "_Slicer";
            var slicer = RequireSuccess(slicers.CreateSlicer(batch, pivotName, hierarchy, slicerName, sheetName, "H3"));
            Assert.Equal([pivotName], slicer.ConnectedPivotTables);
            Assert.NotEmpty(slicer.AvailableItems);
            RequireSuccess(slicers.DeleteSlicer(batch, slicerName));

            var charts = _fixture.CreateCommands<IChartCommands>();
            var chart = RequireSuccess(charts.CreateFromPivotTable(
                batch, pivotName, sheetName, ChartType.ColumnClustered, chartName: connectionName + "_Chart"));
            RequireSuccess(charts.Delete(batch, chart.ChartName));

            var removed = RequireSuccess(pivots.RemoveField(batch, pivotName, hierarchy));
            Assert.Equal(PivotFieldArea.Hidden, removed.Area);
        }
        finally
        {
            _fixture.Send("sheet.delete", new { sheetName });
            _fixture.ForgetSheet(sheetName);
        }
    }
}

using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryManualTableTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string OriginalMCode = """
        let
            Source = #table(
                {"Column1", "Column2"},
                {{"A", "B"}, {"C", "D"}}
            )
        in
            Source
        """;

    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);
    private readonly IPersistentTableCommands _tables =
        ServiceCommandProxy.Create<IPersistentTableCommands>(fixture);

    [Theory]
    [InlineData(PowerQueryLoadMode.LoadToTable)]
    [InlineData(PowerQueryLoadMode.LoadToDataModel)]
    [InlineData(PowerQueryLoadMode.ConnectionOnly)]
    public void RefreshAll_OrdinaryTableBeforeQueryDestination_DoesNotBlockRefresh(PowerQueryLoadMode mode)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var suffix = Guid.NewGuid().ToString("N")[..8];
        var tableName = $"RefreshSource_{suffix}";
        var queryName = $"RefreshTarget_{suffix}";
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B3", [["Key", "Value"], [1, 10], [2, 20]]));
        RequireSuccess(_tables.Create(batch, sheet, tableName, "A1:B3"));
        var code = $"let Source = Excel.CurrentWorkbook(){{[Name=\"{tableName}\"]}}[Content], " +
            "Typed = Table.TransformColumnTypes(Source, {{\"Key\", Int64.Type}, {\"Value\", Int64.Type}}) in Typed";
        var worksheetLoad = mode == PowerQueryLoadMode.LoadToTable;
        RequireSuccess(_queries.Create(batch, queryName, code, mode,
            worksheetLoad ? sheet : null, worksheetLoad ? "D1" : null));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        if (worksheetLoad)
        {
            PowerQueryStateAssertions.AssertRows([["Key", "Value"], [1, 10], [2, 20]],
                RequireSuccess(_commands.GetValues(batch, sheet, "D1:E3")).Values);
        }

        RequireSuccess(_commands.SetValues(batch, sheet, "A2:B3", [[7, 70], [9, 90]],
            overwritePolicy: OverwritePolicy.Allow));
        var result = RequireSuccess(_queries.RefreshAll(batch, refreshPivotTables: false));

        Assert.Empty(result.FailedQueries);
        if (mode == PowerQueryLoadMode.ConnectionOnly)
        {
            Assert.Empty(result.RefreshedQueries);
            Assert.Equal(queryName, Assert.Single(result.SkippedQueries).QueryName);
        }
        else
        {
            Assert.Equal(queryName, Assert.Single(result.RefreshedQueries));
            Assert.Empty(result.SkippedQueries);
        }
        var view = RequireSuccess(_queries.View(batch, queryName));
        Assert.Equal(mode, view.LoadMode);
        Assert.Equal(code, view.MCode);
        Assert.Equal(worksheetLoad ? sheet : null, view.TargetSheet);
        if (worksheetLoad)
        {
            PowerQueryStateAssertions.AssertRows([["Key", "Value"], [7, 70], [9, 90]],
                RequireSuccess(_commands.GetValues(batch, sheet, "D1:E3")).Values);
        }
        if (mode == PowerQueryLoadMode.LoadToDataModel)
        {
            var model = _fixture.CreateCommands<IDataModelCommands>();
            var rows = RequireSuccess(model.Evaluate(batch, $"EVALUATE '{queryName}'"));
            Assert.Equal([$"{queryName}[Key]", $"{queryName}[Value]"], rows.Columns);
            Assert.Equal(2, rows.RowCount);
            Assert.Equal(2, rows.ColumnCount);
            PowerQueryStateAssertions.AssertRows([[7, 70], [9, 90]], rows.Rows);
        }
        PowerQueryStateAssertions.AssertRows([["Key", "Value"], [7, 70], [9, 90]],
            RequireSuccess(_commands.GetValues(batch, sheet, "A1:B3")).Values);
        var table = RequireSuccess(_tables.Read(batch, tableName)).Table;
        Assert.NotNull(table);
        Assert.Equal(tableName, table.Name);
        Assert.Equal(sheet, table.SheetName);
        Assert.Equal("$A$1:$B$3", table.Range);
    }

    [Fact]
    public void Update_WorkbookWithManualTable_UpdatesQuerySuccessfully()
    {
        var state = ArrangeManualTableAndQuery();
        var listed = RequireSuccess(_queries.List(_fixture.BatchToken));
        var query = Assert.Single(listed.Queries);
        Assert.DoesNotContain(listed.Queries,
            candidate => candidate.Name.StartsWith("Error Query", StringComparison.Ordinal));
        Assert.Equal(state.QueryName, query.Name);
        Assert.NotEmpty(query.FormulaPreview);
        Assert.DoesNotContain("Error:", query.FormulaPreview);
        Assert.True(query.IsConnectionOnly);
        const string updatedMCode = """
            let
                Source = #table(
                    {"NewCol1", "NewCol2", "NewCol3"},
                    {{1, 2, 3}, {4, 5, 6}}
                )
            in
                Source
            """;

        RequireSuccess(_queries.Update(
            _fixture.BatchToken,
            state.QueryName,
            updatedMCode));

        var result = RequireSuccess(_queries.View(_fixture.BatchToken, state.QueryName));
        Assert.True(result.Success, $"View after update failed: {result.ErrorMessage}");
        Assert.Contains("NewCol1", result.MCode);
        Assert.Contains("NewCol2", result.MCode);
        Assert.Contains("NewCol3", result.MCode);
        Assert.DoesNotContain("Column1", result.MCode);
        AssertPreserved(state, updatedMCode, ["NewCol1", "NewCol2", "NewCol3"],
            [[1, 2, 3], [4, 5, 6]]);
    }

    private ManualState ArrangeManualTableAndQuery()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var suffix = Guid.NewGuid().ToString("N")[..8];
        var tableName = $"ManualTable_{suffix}";
        var queryName = $"TestQuery_{suffix}";
        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A1:B3",
            [["Header1", "Header2"], ["Data1", "Data2"], ["Data3", "Data4"]]));
        RequireSuccess(_tables.Create(
            batch,
            sheetName,
            tableName,
            "A1:B3",
            true,
            "TableStyleMedium2"));
        RequireSuccess(_queries.Create(
            batch,
            queryName,
            OriginalMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        var state = new ManualState(queryName, tableName, sheetName,
            JsonSerializer.Serialize(RequireSuccess(_tables.Read(batch, tableName))));
        AssertPreserved(state, OriginalMCode, ["Column1", "Column2"], [["A", "B"], ["C", "D"]]);
        return state;
    }

    private void AssertPreserved(ManualState state, string code, string[] columns, object[][] rows)
    {
        PowerQueryStateAssertions.AssertStored(_fixture, state.QueryName, code,
            PowerQueryLoadMode.ConnectionOnly, null, columns, rows);
        var result = RequireSuccess(_queries.Evaluate(_fixture.BatchToken, code));
        Assert.Equal(code, result.MCode);
        Assert.Equal(columns, result.Columns);
        Assert.Equal(columns.Length, result.ColumnCount);
        Assert.Equal(rows.Length, result.RowCount);
        PowerQueryStateAssertions.AssertRows(rows, result.Rows);
        PowerQueryStateAssertions.AssertRows(
            [["Header1", "Header2"], ["Data1", "Data2"], ["Data3", "Data4"]],
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, state.SheetName, "A1:B3")).Values);
        Assert.Equal(state.TableMetadata,
            JsonSerializer.Serialize(RequireSuccess(_tables.Read(_fixture.BatchToken, state.TableName))));
        PowerQueryStateAssertions.AssertStored(_fixture, state.QueryName, code,
            PowerQueryLoadMode.ConnectionOnly, null, columns, rows);
    }

    private sealed record ManualState(string QueryName, string TableName, string SheetName, string TableMetadata);
}

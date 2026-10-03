using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
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

    [Fact]
    public void List_WorkbookWithManualTable_ReturnsOnlyQueries()
    {
        var state = ArrangeManualTableAndQuery();

        var result = RequireSuccess(_queries.List(_fixture.BatchToken));

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        var query = Assert.Single(result.Queries);
        Assert.DoesNotContain(
            result.Queries,
            candidate => candidate.Name.StartsWith(
                "Error Query",
                StringComparison.Ordinal));
        Assert.Equal(state.QueryName, query.Name);
        Assert.NotEmpty(query.FormulaPreview);
        Assert.DoesNotContain("Error:", query.FormulaPreview);
        Assert.True(query.IsConnectionOnly);
        AssertPreserved(state, OriginalMCode, ["Column1", "Column2"], [["A", "B"], ["C", "D"]]);
    }

    [Fact]
    public void View_WorkbookWithManualTable_ReturnsQueryDetails()
    {
        var state = ArrangeManualTableAndQuery();

        var result = RequireSuccess(_queries.View(_fixture.BatchToken, state.QueryName));

        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(state.QueryName, result.QueryName);
        Assert.NotEmpty(result.MCode);
        Assert.Contains("Source = #table", result.MCode);
        Assert.True(result.IsConnectionOnly);
        AssertPreserved(state, OriginalMCode, ["Column1", "Column2"], [["A", "B"], ["C", "D"]]);
    }

    [Fact]
    public void Update_WorkbookWithManualTable_UpdatesQuerySuccessfully()
    {
        var state = ArrangeManualTableAndQuery();
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

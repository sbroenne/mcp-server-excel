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
        var queryName = ArrangeManualTableAndQuery();

        var result = _queries.List(_fixture.BatchToken);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        var query = Assert.Single(result.Queries);
        Assert.DoesNotContain(
            result.Queries,
            candidate => candidate.Name.StartsWith(
                "Error Query",
                StringComparison.Ordinal));
        Assert.Equal(queryName, query.Name);
        Assert.NotEmpty(query.FormulaPreview);
        Assert.DoesNotContain("Error:", query.FormulaPreview);
        Assert.True(query.IsConnectionOnly);
    }

    [Fact]
    public void View_WorkbookWithManualTable_ReturnsQueryDetails()
    {
        var queryName = ArrangeManualTableAndQuery();

        var result = _queries.View(_fixture.BatchToken, queryName);

        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(queryName, result.QueryName);
        Assert.NotEmpty(result.MCode);
        Assert.Contains("Source = #table", result.MCode);
        Assert.True(result.IsConnectionOnly);
    }

    [Fact]
    public void Update_WorkbookWithManualTable_UpdatesQuerySuccessfully()
    {
        var queryName = ArrangeManualTableAndQuery();
        const string updatedMCode = """
            let
                Source = #table(
                    {"NewCol1", "NewCol2", "NewCol3"},
                    {{1, 2, 3}, {4, 5, 6}}
                )
            in
                Source
            """;

        _queries.Update(
            _fixture.BatchToken,
            queryName,
            updatedMCode);

        var result = _queries.View(_fixture.BatchToken, queryName);
        Assert.True(result.Success, $"View after update failed: {result.ErrorMessage}");
        Assert.Contains("NewCol1", result.MCode);
        Assert.Contains("NewCol2", result.MCode);
        Assert.Contains("NewCol3", result.MCode);
        Assert.DoesNotContain("Column1", result.MCode);
    }

    private string ArrangeManualTableAndQuery()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var suffix = Guid.NewGuid().ToString("N")[..8];
        var tableName = $"ManualTable_{suffix}";
        var queryName = $"TestQuery_{suffix}";
        _commands.SetValues(
            batch,
            sheetName,
            "A1:B3",
            [["Header1", "Header2"], ["Data1", "Data2"], ["Data3", "Data4"]]);
        _tables.Create(
            batch,
            sheetName,
            tableName,
            "A1:B3",
            true,
            "TableStyleMedium2");
        _queries.Create(
            batch,
            queryName,
            OriginalMCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        return queryName;
    }
}

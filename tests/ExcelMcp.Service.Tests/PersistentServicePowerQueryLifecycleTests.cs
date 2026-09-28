using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServicePowerQueryLifecycleTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string TableMCode = """
        let
            Source = #table(
                {"Column1", "Column2", "Column3"},
                {
                    {"Value1", "Value2", "Value3"},
                    {"A", "B", "C"},
                    {"X", "Y", "Z"}
                }
            )
        in
            Source
        """;

    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);
    private readonly IDataModelCommands _dataModel =
        ServiceCommandProxy.Create<IDataModelCommands>(fixture);
    private readonly IConnectionCommands _connections =
        ServiceCommandProxy.Create<IConnectionCommands>(fixture);

    [Fact]
    public void Import_ValidMCode_ReturnsSuccess()
    {
        const string queryName = "TestQuery";
        var batch = _fixture.BatchToken;

        _queries.Create(batch, queryName, TableMCode);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var result = _queries.List(batch);
        Assert.Contains(result.Queries, query => query.Name == queryName);
    }

    [Fact]
    public void Update_ExistingQuery_ReturnsSuccess()
    {
        var queryName = "PQ_Update_" + Guid.NewGuid().ToString("N")[..8];
        const string updatedMCode = """
            let
                UpdatedSource = 1
            in
                UpdatedSource
            """;
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, TableMCode);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        _queries.Update(batch, queryName, updatedMCode);
    }

    [Fact]
    public void Update_ExistingQuery_ReplacesNotMergesMCode()
    {
        var queryName = "PQ_ReplaceTest_" + Guid.NewGuid().ToString("N")[..8];
        const string originalMCode = """
            let
                OriginalSource = "ORIGINAL_MARKER",
                OriginalStep = "Should be completely removed"
            in
                OriginalSource
            """;
        const string newMCode = """
            let
                NewSource = "NEW_MARKER",
                NewStep = "Should be the only content"
            in
                NewSource
            """;
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, originalMCode);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        _queries.Update(batch, queryName, newMCode);
        var result = _queries.View(batch, queryName);

        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Contains("NEW_MARKER", result.MCode);
        Assert.Contains("NewSource", result.MCode);
        Assert.Contains("Should be the only content", result.MCode);
        Assert.DoesNotContain("ORIGINAL_MARKER", result.MCode);
        Assert.DoesNotContain("OriginalSource", result.MCode);
        Assert.DoesNotContain("Should be completely removed", result.MCode);
        Assert.Single(Regex.Matches(result.MCode, @"\blet\b").Cast<Match>());
        Assert.Single(Regex.Matches(result.MCode, @"\bin\b").Cast<Match>());
    }

    [Fact]
    public void Update_MultipleSequentialUpdates_EachReplacesCompletely()
    {
        var queryName = "PQ_MultiUpdate_" + Guid.NewGuid().ToString("N")[..8];
        const string version1 = "let V1 = \"VERSION_1\" in V1";
        const string version2 = "let V2 = \"VERSION_2\" in V2";
        const string version3 = "let V3 = \"VERSION_3\" in V3";
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, version1);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        _queries.Update(batch, queryName, version2);
        _queries.Update(batch, queryName, version3);
        var result = _queries.View(batch, queryName);

        Assert.Contains("VERSION_3", result.MCode);
        Assert.DoesNotContain("VERSION_1", result.MCode);
        Assert.DoesNotContain("VERSION_2", result.MCode);
        Assert.Single(Regex.Matches(result.MCode, @"\blet\b").Cast<Match>());
        Assert.Single(Regex.Matches(result.MCode, @"\bin\b").Cast<Match>());
    }

    [Fact]
    public void Update_InvalidMCodeSyntax_AcceptedByExcel()
    {
        var queryName = $"PQ_Invalid_{Guid.NewGuid():N}"[..20];
        const string validMCode = """
            let
                Source = #table({"A"}, {{1}})
            in
                Source
            """;
        const string invalidMCode = """
            let
                Source = this is not valid M code syntax!!!
            in
                Source
            """;
        var batch = _fixture.BatchToken;
        _queries.Create(
            batch,
            queryName,
            validMCode,
            Sbroenne.ExcelMcp.Core.Models.PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        _queries.Update(batch, queryName, invalidMCode);

        var result = _queries.View(batch, queryName);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains("not valid", result.MCode);
    }

    [Fact]
    public void Update_ValidMCode_Succeeds()
    {
        var queryName = $"PQ_Valid_{Guid.NewGuid():N}"[..20];
        const string initialMCode =
            "let Source = #table({\"A\"}, {{1}}) in Source";
        const string updatedMCode =
            "let Source = #table({\"A\", \"B\"}, {{1, 2}}) in Source";
        var batch = _fixture.BatchToken;
        _queries.Create(
            batch,
            queryName,
            initialMCode,
            Sbroenne.ExcelMcp.Core.Models.PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        _queries.Update(batch, queryName, updatedMCode);

        var result = _queries.View(batch, queryName);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains("B", result.MCode);
    }

    [Fact]
    public void Update_NonExistentQuery_ThrowsWithMeaningfulMessage()
    {
        var queryName = $"PQ_Missing_{Guid.NewGuid():N}"[..20];

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Update(
                _fixture.BatchToken,
                queryName,
                "let Source = 1 in Source"));

        Assert.Contains(queryName, exception.Message);
    }

    [Fact]
    public void Delete_ExistingQuery_ReturnsSuccess()
    {
        var queryName = "PQ_Delete_" + Guid.NewGuid().ToString("N")[..8];
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, TableMCode);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        _queries.Delete(batch, queryName);
        _fixture.ForgetPowerQuery(queryName);
    }

    [Fact]
    public void Create_DuplicateQueryName_ReturnsError()
    {
        const string queryName = "TestQuery";
        var batch = _fixture.BatchToken;
        _queries.Create(batch, queryName, TableMCode);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Create(batch, queryName, TableMCode));

        Assert.Contains(
            "already exists",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains(queryName, exception.Message);
        var result = _queries.View(batch, queryName);
        Assert.True(result.Success);
        Assert.NotEmpty(result.MCode);
    }
}

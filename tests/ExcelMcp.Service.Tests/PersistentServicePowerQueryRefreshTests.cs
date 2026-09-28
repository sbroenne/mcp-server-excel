using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServicePowerQueryRefreshTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string ValidMCode =
        "let Source = #table({\"X\"}, {{1}}) in Source";

    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);

    [Fact]
    public void Refresh_WorksheetQueryWithInvalidMCode_ThrowsError()
    {
        var queryName = CreateWorksheetQuery("BrokenWorksheet");
        _queries.Update(
            _fixture.BatchToken,
            queryName,
            "let Source = NonExistentFunction() in Source",
            refresh: false);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Refresh(
                _fixture.BatchToken,
                queryName,
                TimeSpan.FromMinutes(1)));

        Assert.Contains(
            "Expression.Error",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Refresh_QueryReferencingNonExistentTable_ThrowsError()
    {
        var queryName = CreateWorksheetQuery("TableRef");
        const string invalidMCode = """
            let
                Source = Excel.CurrentWorkbook(){[Name="NonExistentTable"]}[Content]
            in
                Source
            """;
        _queries.Update(
            _fixture.BatchToken,
            queryName,
            invalidMCode,
            refresh: false);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Refresh(
                _fixture.BatchToken,
                queryName,
                TimeSpan.FromMinutes(1)));

        Assert.True(
            exception.Message.Contains(
                "Expression.Error",
                StringComparison.OrdinalIgnoreCase)
            || exception.Message.Contains(
                "DataSource.Error",
                StringComparison.OrdinalIgnoreCase)
            || exception.Message.Contains(
                "didn't find",
                StringComparison.OrdinalIgnoreCase),
            $"Expected Power Query error but got: {exception.Message}");
    }

    [Fact]
    public void Refresh_QueryWithSyntheticFirewallError_ReturnsUsefulPublicError()
    {
        var queryName = CreateWorksheetQuery("SyntheticFirewall");
        const string firewallMCode = """
            let
                Root = error Error.Record(
                    "Formula.Firewall",
                    "Query 'ConfigData' (step 'Root') references other queries or steps, so it may not directly access a data source.",
                    null)
            in
                Root
            """;
        _queries.Update(
            _fixture.BatchToken,
            queryName,
            firewallMCode,
            refresh: false);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Refresh(
                _fixture.BatchToken,
                queryName,
                TimeSpan.FromMinutes(1)));

        Assert.Contains(
            "PowerQueryCommandException",
            exception.Message,
            StringComparison.Ordinal);
        Assert.Contains(
            "Formula.Firewall",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains(
            "may not directly access a data source",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Refresh_ValidWorksheetQuery_Succeeds()
    {
        var queryName = CreateWorksheetQuery("ValidWorksheet");

        var result = _queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(1));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
    }

    [Fact]
    public void Refresh_ZeroTimeout_UsesDefaultAndSucceeds()
    {
        var queryName = CreateWorksheetQuery("ZeroTimeout");

        var result = _queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.Zero);

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
    }

    [Fact]
    public void Refresh_ConnectionOnlyQuery_ThrowsBecauseNoRefreshMechanism()
    {
        var queryName = UniqueName("ConnectionOnly");
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            ValidMCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Refresh(
                _fixture.BatchToken,
                queryName,
                TimeSpan.FromMinutes(1)));

        Assert.Contains(
            "Could not find connection or table",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RefreshAll_ZeroTimeout_UsesDefaultAndSucceeds()
    {
        CreateWorksheetQuery("RefreshAllZero");

        var result = _queries.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.Zero);

        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
    }

    private string CreateWorksheetQuery(string prefix)
    {
        var queryName = UniqueName(prefix);
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            ValidMCode,
            PowerQueryLoadMode.LoadToTable,
            queryName);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        return queryName;
    }

    private static string UniqueName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];
}

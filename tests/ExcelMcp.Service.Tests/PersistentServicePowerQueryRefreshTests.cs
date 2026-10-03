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
public sealed partial class PersistentServicePowerQueryRefreshTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private const string ValidMCode =
        "let Source = #table({\"X\"}, {{1}}) in Source";

    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);
    private readonly Dictionary<string, string> _storedSources = new(StringComparer.Ordinal);

    [Fact]
    public void Refresh_WorksheetQueryWithInvalidMCode_ThrowsError()
    {
        var queryName = CreateWorksheetQuery("BrokenWorksheet");
        var guard = CreateRefreshGuard();
        StageSource(queryName, "let Source = NonExistentFunction() in Source");
        AssertWorksheetValue(queryName, 1);
        var before = SnapshotQueries();

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Refresh(
                _fixture.BatchToken,
                queryName,
                TimeSpan.FromMinutes(1)));

        Assert.Contains(
            "Expression.Error",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        AssertWorksheetValue(queryName, 1);
        Assert.Equal(before, SnapshotQueries());
        AssertRefreshGuard(guard);
    }

    [Fact]
    public void Refresh_QueryReferencingNonExistentTable_ThrowsError()
    {
        var queryName = CreateWorksheetQuery("TableRef");
        var guard = CreateRefreshGuard();
        const string invalidMCode = """
            let
                Source = Excel.CurrentWorkbook(){[Name="NonExistentTable"]}[Content]
            in
                Source
            """;
        StageSource(queryName, invalidMCode);
        AssertWorksheetValue(queryName, 1);
        var before = SnapshotQueries();

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
        AssertWorksheetValue(queryName, 1);
        Assert.Contains("NonExistentTable", exception.Message);
        Assert.Equal(before, SnapshotQueries());
        AssertRefreshGuard(guard);
    }

    [Fact]
    public void Refresh_QueryWithSyntheticFirewallError_ReturnsUsefulPublicError()
    {
        var queryName = CreateWorksheetQuery("SyntheticFirewall");
        var guard = CreateRefreshGuard();
        const string firewallMCode = """
            let
                Root = error Error.Record(
                    "Formula.Firewall",
                    "Query 'ConfigData' (step 'Root') references other queries or steps, so it may not directly access a data source.",
                    null)
            in
                Root
            """;
        StageSource(queryName, firewallMCode);
        AssertWorksheetValue(queryName, 1);
        var before = SnapshotQueries();

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
        AssertWorksheetValue(queryName, 1);
        Assert.Equal(before, SnapshotQueries());
        AssertRefreshGuard(guard);
    }

    [Fact]
    public void Refresh_ValidWorksheetQuery_Succeeds()
    {
        var queryName = CreateWorksheetQuery("ValidWorksheet");
        StageWorksheetUpdate(queryName, 42);

        var result = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(1)));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
        AssertRefreshMetadata(result, queryName, queryName);
        AssertWorksheetValue(queryName, 42);
    }

    [Fact]
    public void Refresh_ZeroTimeout_UsesDefaultAndSucceeds()
    {
        var queryName = CreateWorksheetQuery("ZeroTimeout");
        StageWorksheetUpdate(queryName, 73);

        var result = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.Zero));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
        AssertRefreshMetadata(result, queryName, queryName);
        AssertWorksheetValue(queryName, 73);
    }

    [Fact]
    public async Task Refresh_ConnectionOnlyQuery_ReturnsCategorizedPrerequisite()
    {
        var queryName = UniqueName("ConnectionOnly");
        var guard = CreateRefreshGuard();
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            ValidMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, ValidMCode,
            PowerQueryLoadMode.ConnectionOnly, null, ["X"], [[1]]);
        var before = SnapshotQueries();

        var response = await _fixture.SendForFailureAsync(
            "powerquery.refresh",
            new { queryName, timeout = 60 });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal("Prerequisite", response.ErrorCategory);
        Assert.Contains(
            "Could not find connection or table",
            response.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, SnapshotQueries());
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, ValidMCode,
            PowerQueryLoadMode.ConnectionOnly, null, ["X"], [[1]]);
        AssertRefreshGuard(guard);
    }

    [Fact]
    public void RefreshAll_ZeroTimeout_UsesDefaultAndSucceeds()
    {
        var first = CreateWorksheetQuery("RefreshAllZero");
        var second = CreateWorksheetQuery("RefreshAllOther");
        StageWorksheetUpdate(first, 41);
        StageWorksheetUpdate(second, 82);

        var result = RequireSuccess(_queries.RefreshAll(
            _fixture.BatchToken,
            TimeSpan.Zero));

        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
        AssertWorksheetValue(first, 41);
        AssertWorksheetValue(second, 82);
    }

    [Fact]
    public void RefreshAll_DefinitionOnlyStage_ReportsNamedRefreshRecoveryAndRefreshesLoadedDependent()
    {
        var stageName = UniqueName("Stage");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken, stageName, ValidMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(stageName);
        var loadedName = UniqueName("Loaded");
        RequireSuccess(_queries.Create(
            _fixture.BatchToken, loadedName, stageName,
            PowerQueryLoadMode.LoadToTable, loadedName));
        _fixture.RegisterPowerQueryForCleanup(loadedName);
        _fixture.RegisterSheetForCleanup(loadedName);
        PowerQueryStateAssertions.AssertStored(_fixture, loadedName, stageName,
            PowerQueryLoadMode.LoadToTable, loadedName, ["X"], [[1]]);
        const string changedCode = "let Source = #table({\"X\"}, {{42}}) in Source";
        RequireSuccess(_queries.Update(
            _fixture.BatchToken, stageName,
            changedCode, refresh: false));
        var pending = RequireSuccess(_commands.GetValues(_fixture.BatchToken, loadedName, "A1:A2"));
        Assert.Equal("X", pending.Values[0][0]);
        Assert.Equal(1d, Convert.ToDouble(pending.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));

        var error = Assert.Throws<InvalidOperationException>(() =>
            _queries.RefreshAll(_fixture.BatchToken, TimeSpan.FromMinutes(1)));
        Assert.Contains(stageName, error.Message, StringComparison.Ordinal);
        Assert.Contains("loaded dependent", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("powerquery refresh", error.Message, StringComparison.OrdinalIgnoreCase);
        PowerQueryStateAssertions.AssertStored(_fixture, loadedName, stageName,
            PowerQueryLoadMode.LoadToTable, loadedName, ["X"], [[42]]);
        PowerQueryStateAssertions.AssertStored(_fixture, stageName, changedCode,
            PowerQueryLoadMode.ConnectionOnly, null, ["X"], [[42]]);
        var recovery = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken, loadedName, TimeSpan.FromMinutes(1)));
        AssertRefreshMetadata(recovery, loadedName, loadedName);
        PowerQueryStateAssertions.AssertStored(_fixture, loadedName, stageName,
            PowerQueryLoadMode.LoadToTable, loadedName, ["X"], [[42]]);
    }

    private string CreateWorksheetQuery(string prefix)
    {
        var queryName = UniqueName(prefix);
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            ValidMCode,
            PowerQueryLoadMode.LoadToTable,
            queryName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        _storedSources.Add(queryName, ValidMCode);
        AssertWorksheetValue(queryName, 1);
        return queryName;
    }

    private static string UniqueName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];

    private void StageWorksheetUpdate(string queryName, int newValue)
    {
        AssertWorksheetValue(queryName, 1);
        StageSource(queryName,
            $"let Source = #table({{\"X\"}}, {{{{{newValue}}}}}) in Source");
        AssertWorksheetValue(queryName, 1);
    }

    private void AssertWorksheetValue(string queryName, int expected)
    {
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, _storedSources[queryName],
            PowerQueryLoadMode.LoadToTable, queryName, ["X"], [[expected]]);
    }

    private void StageSource(string name, string code)
    {
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, code, refresh: false));
        _storedSources[name] = code;
    }

    private string SnapshotQueries() =>
        JsonSerializer.Serialize(RequireSuccess(_queries.List(_fixture.BatchToken)).Queries);

    private const string RefreshGuardCode =
        "let Source = #table(type table [Guard = Int64.Type], {{47}, {83}}) in Source";

    private string CreateRefreshGuard()
    {
        var name = UniqueName("RefreshGuard");
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, RefreshGuardCode,
            PowerQueryLoadMode.LoadToBoth, name));
        _fixture.RegisterPowerQueryForCleanup(name);
        _fixture.RegisterSheetForCleanup(name);
        AssertRefreshGuard(name);
        return name;
    }

    private void AssertRefreshGuard(string name) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, RefreshGuardCode,
            PowerQueryLoadMode.LoadToBoth, name, ["Guard"], [[47], [83]]);

    private static void AssertRefreshMetadata(PowerQueryRefreshResult result, string name, string? sheet)
    {
        Assert.Equal(name, result.QueryName);
        Assert.Equal(sheet, result.LoadedToSheet);
        Assert.False(result.IsConnectionOnly);
        Assert.False(result.HasErrors);
        Assert.Empty(result.ErrorMessages);
        Assert.NotEqual(default, result.RefreshTime);
    }
}

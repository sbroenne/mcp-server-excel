using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Collection("ServiceWorkflow")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public class PersistentServicePowerQueryWorksheetCleanupTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();
    private readonly IConnectionCommands _connections =
        fixture.CreateCommands<IConnectionCommands>();
    private readonly ITableCommands _tables =
        fixture.CreateCommands<ITableCommands>();

    [Fact]
    public void Create_LoadToTable_CreatesProperlyNamedConnection() =>
        AssertLoaded(CreateQuery(PowerQueryLoadMode.LoadToTable, [11, 23, 37]));

    [Fact]
    public void Create_MultipleLoadToTable_NoOrphanedConnections()
    {
        var first = CreateQuery(PowerQueryLoadMode.LoadToTable, [11, 12]);
        var second = CreateQuery(PowerQueryLoadMode.LoadToTable, [23, 24]);
        var third = CreateQuery(PowerQueryLoadMode.LoadToTable, [37, 38]);
        AssertLoaded(first);
        AssertLoaded(second);
        AssertLoaded(third);
        var connections = RequireSuccess(_connections.List(_fixture.BatchToken)).Connections;
        Assert.Equal(3, connections.Count(connection => connection.IsPowerQuery));
        Assert.DoesNotContain(connections, connection => connection.Name == "Connection" ||
            connection.Name == "Connection1");
    }

    [Fact]
    public void Delete_OneOfMultipleLoadedQueries_OnlyRemovesItsOwnResources() =>
        AssertRemoval(PowerQueryLoadMode.LoadToTable, delete: true);

    [Fact]
    public void Delete_ConnectionOnly_CleanSlate() =>
        AssertRemoval(PowerQueryLoadMode.ConnectionOnly, delete: true);

    [Fact]
    public void Unload_LoadedToWorksheet_RemovesTableAndConnectionKeepsQuery() =>
        AssertRemoval(PowerQueryLoadMode.LoadToTable, delete: false);

    [Fact]
    public void Delete_ExistingQuery_VerifiesCleanSlate()
    {
        var name = UniqueName();
        const string code =
            "let Source = #table({\"First\", \"Second\", \"Third\"}, " +
            "{{\"A\", \"B\", \"C\"}, {\"D\", \"E\", \"F\"}, {\"G\", \"H\", \"I\"}}) in Source";
        var state = new QueryState(name, code, PowerQueryLoadMode.LoadToTable, name,
            ["First", "Second", "Third"], [["A", "B", "C"], ["D", "E", "F"], ["G", "H", "I"]]);
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, code,
            PowerQueryLoadMode.LoadToTable, name));
        _fixture.RegisterPowerQueryForCleanup(name);
        _fixture.RegisterSheetForCleanup(name);
        AssertLoaded(state);
        var neighbor = CreateQuery(PowerQueryLoadMode.LoadToBoth, [47, 83], name + "A");
        RequireSuccess(_queries.Delete(_fixture.BatchToken, name));
        _fixture.ForgetPowerQuery(name);
        AssertRemoved(state, delete: true);
        AssertLoaded(neighbor);
    }

    [Fact]
    public void LoadTo_ExistingConnectionOnlyQuery_CreatesProperlyNamedConnection() =>
        AssertTransition(PowerQueryLoadMode.ConnectionOnly, PowerQueryLoadMode.LoadToTable);

    [Fact]
    public void Refresh_LoadedQuery_MaintainsProperConnectionNaming()
    {
        var state = CreateQuery(PowerQueryLoadMode.LoadToTable, [17, 29]);
        var neighbor = CreateQuery(PowerQueryLoadMode.LoadToBoth, [47, 83], state.Name + "A");
        var updated = state with { Code = Code("Val", [61, 73, 89]), Rows = [[61], [73], [89]] };
        RequireSuccess(_queries.Update(_fixture.BatchToken, state.Name, updated.Code, refresh: false));
        AssertLoaded(state with { Code = updated.Code });
        RequireSuccess(_queries.Refresh(_fixture.BatchToken, state.Name, TimeSpan.FromMinutes(2)));
        AssertLoaded(updated);
        AssertLoaded(neighbor);
    }

    [Fact]
    public void Update_LoadedQuery_MaintainsProperConnectionNaming()
    {
        var state = CreateQuery(PowerQueryLoadMode.LoadToTable, [17, 29]);
        var neighbor = CreateQuery(PowerQueryLoadMode.LoadToBoth, [47, 83], state.Name + "A");
        var updated = state with
        {
            Code = Code("NewVal", [61, 73, 89]),
            Columns = ["NewVal"],
            Rows = [[61], [73], [89]]
        };
        RequireSuccess(_queries.Update(_fixture.BatchToken, state.Name, updated.Code));
        AssertLoaded(updated);
        AssertLoaded(neighbor);
        RequireSuccess(_queries.Delete(_fixture.BatchToken, state.Name));
        _fixture.ForgetPowerQuery(state.Name);
        AssertRemoved(updated, delete: true);
        AssertLoaded(neighbor);
    }

    [Fact]
    public void LoadTo_ConnectionOnlyToBoth_CreatesDualConnectionsProperly() =>
        AssertTransition(PowerQueryLoadMode.ConnectionOnly, PowerQueryLoadMode.LoadToBoth);

    [Fact]
    public void Create_LoadToBoth_ExactlyTwoConnectionsWithProperNaming()
    {
        var state = CreateQuery(PowerQueryLoadMode.LoadToBoth, [17, 29]);
        var neighbor = CreateQuery(PowerQueryLoadMode.LoadToTable, [47, 83], state.Name + "A");
        AssertLoaded(state);
        RequireSuccess(_queries.Delete(_fixture.BatchToken, state.Name));
        _fixture.ForgetPowerQuery(state.Name);
        AssertRemoved(state, delete: true);
        AssertLoaded(neighbor);
    }

    [Fact]
    public void UnloadThenReload_NoOrphanedConnections()
    {
        var state = CreateQuery(PowerQueryLoadMode.LoadToTable, [17, 29]);
        var neighbor = CreateQuery(PowerQueryLoadMode.LoadToBoth, [47, 83], state.Name + "A");
        RequireSuccess(_queries.Unload(_fixture.BatchToken, state.Name));
        AssertRemoved(state, delete: false);
        AssertLoaded(neighbor);
        var reloaded = ChangeMode(state, PowerQueryLoadMode.LoadToTable, state.Name + "New");
        AssertLoaded(reloaded);
        AssertClearedSheet(state);
        AssertLoaded(neighbor);
        RequireSuccess(_queries.Delete(_fixture.BatchToken, state.Name));
        _fixture.ForgetPowerQuery(state.Name);
        AssertRemoved(reloaded, delete: true);
        AssertLoaded(neighbor);
    }

    [Fact]
    public void LoadTo_LoadedToTable_ThenConnectionOnly_RemovesTableAndConnection() =>
        AssertTransition(PowerQueryLoadMode.LoadToTable, PowerQueryLoadMode.ConnectionOnly);

    [Fact]
    public void LoadTo_LoadedToTable_ThenLoadToDataModel_RemovesTableAddsDataModel() =>
        AssertTransition(PowerQueryLoadMode.LoadToTable, PowerQueryLoadMode.LoadToDataModel);

    [Fact]
    public void LoadTo_LoadedToDataModel_ThenLoadToTable_RemovesDataModelAddsTable() =>
        AssertTransition(PowerQueryLoadMode.LoadToDataModel, PowerQueryLoadMode.LoadToTable);

    [Fact]
    public void LoadTo_LoadedToBoth_ThenConnectionOnly_RemovesBothDestinations() =>
        AssertTransition(PowerQueryLoadMode.LoadToBoth, PowerQueryLoadMode.ConnectionOnly);

    private void AssertRemoval(PowerQueryLoadMode mode, bool delete)
    {
        var state = CreateQuery(mode, [17, 29]);
        var neighbor = CreateQuery(PowerQueryLoadMode.LoadToBoth, [47, 83], state.Name + "A");
        if (delete)
        {
            RequireSuccess(_queries.Delete(_fixture.BatchToken, state.Name));
            _fixture.ForgetPowerQuery(state.Name);
        }
        else { RequireSuccess(_queries.Unload(_fixture.BatchToken, state.Name)); }
        AssertRemoved(state, delete);
        AssertLoaded(neighbor);
    }

    private void AssertTransition(PowerQueryLoadMode before, PowerQueryLoadMode after)
    {
        var state = CreateQuery(before, [17, 29]);
        var neighbor = CreateQuery(PowerQueryLoadMode.LoadToBoth, [47, 83], state.Name + "A");
        var updated = ChangeMode(state, after, state.Name + "New");
        AssertLoaded(updated);
        if (state.Sheet is not null && state.Sheet != updated.Sheet) { AssertClearedSheet(state); }
        AssertLoaded(neighbor);
    }

    private QueryState CreateQuery(PowerQueryLoadMode mode, int[] values, string? name = null)
    {
        name ??= UniqueName();
        var sheet = mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? name : null;
        var state = new QueryState(name, Code("Val", values), mode, sheet, ["Val"],
            values.Select(value => new object[] { value }).ToArray());
        RequireSuccess(_queries.Create(_fixture.BatchToken, name, state.Code, mode, sheet));
        _fixture.RegisterPowerQueryForCleanup(name);
        if (sheet is not null) { _fixture.RegisterSheetForCleanup(sheet); }
        AssertLoaded(state);
        return state;
    }

    private QueryState ChangeMode(QueryState state, PowerQueryLoadMode mode, string sheet)
    {
        var target = mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth ? sheet : null;
        RequireSuccess(_queries.LoadTo(_fixture.BatchToken, state.Name, mode, target));
        if (target is not null) { _fixture.RegisterSheetForCleanup(target); }
        return state with { Mode = mode, Sheet = target };
    }

    private void AssertLoaded(QueryState state)
    {
        PowerQueryStateAssertions.AssertStored(_fixture, state.Name, state.Code,
            state.Mode, state.Sheet, state.Columns, state.Rows);
        var tables = RequireSuccess(_tables.List(_fixture.BatchToken)).Tables;
        if (state.Sheet is not null) { Assert.Single(tables, table => table.Name == state.Name); }
        else { Assert.DoesNotContain(tables, table => table.Name == state.Name); }
        var connections = RequireSuccess(_connections.List(_fixture.BatchToken)).Connections;
        if (state.Mode != PowerQueryLoadMode.ConnectionOnly)
        {
            Assert.Single(connections, connection => connection.Name == $"Query - {state.Name}");
        }
        if (state.Mode == PowerQueryLoadMode.LoadToBoth)
        {
            Assert.Single(connections,
                connection => connection.Name == $"Query - {state.Name} (Data Model)");
        }
    }

    private void AssertRemoved(QueryState state, bool delete)
    {
        if (delete) { PowerQueryStateAssertions.AssertRemoved(_fixture, state.Name); }
        else { AssertLoaded(state with { Mode = PowerQueryLoadMode.ConnectionOnly, Sheet = null }); }
        Assert.DoesNotContain(RequireSuccess(_tables.List(_fixture.BatchToken)).Tables,
            table => table.Name == state.Name);
        if (state.Sheet is not null) { AssertClearedSheet(state); }
    }

    private void AssertClearedSheet(QueryState state) =>
        Assert.All(RequireSuccess(_commands.GetValues(_fixture.BatchToken, state.Sheet!,
            $"A1:{(char)('A' + state.Columns.Length - 1)}{state.Rows.Length + 1}")).Values,
            row => Assert.All(row, Assert.Null));

    private static string Code(string column, int[] values) =>
        $"let Source = #table(type table [{column} = Int64.Type], " +
        $"{{{string.Join(", ", values.Select(value => $"{{{value}}}"))}}}) in Source";

    private static string UniqueName() => "PQ_Clean_" + Guid.NewGuid().ToString("N")[..8];

    private sealed record QueryState(string Name, string Code, PowerQueryLoadMode Mode,
        string? Sheet, string[] Columns, object[][] Rows);
}

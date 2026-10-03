using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryLifecycleTests
{
    private const string CleanupMCode =
        "let Source = #table({\"Val\"}, {{31}, {83}}) in Source";

    [Fact]
    public void Unload_DataModelOnly_RemovesDataModelConnection() =>
        AssertCleanup(PowerQueryLoadMode.LoadToDataModel, delete: false);

    [Fact]
    public void Unload_LoadToBoth_RemovesBothWorksheetAndDataModelConnection() =>
        AssertCleanup(PowerQueryLoadMode.LoadToBoth, delete: false);

    [Fact]
    public void Delete_DataModelOnly_RemovesDataModelConnection() =>
        AssertCleanup(PowerQueryLoadMode.LoadToDataModel, delete: true);

    [Fact]
    public void Delete_LoadToBoth_RemovesBothWorksheetAndDataModelConnection() =>
        AssertCleanup(PowerQueryLoadMode.LoadToBoth, delete: true);

    private void AssertCleanup(PowerQueryLoadMode mode, bool delete)
    {
        var guardName = CreateLoadedGuard();
        var queryName = UniqueCleanupName("PQ_Cleanup");
        var sheetName = mode == PowerQueryLoadMode.LoadToBoth
            ? UniqueCleanupName("CleanupSheet") : null;
        RequireSuccess(_queries.Create(_fixture.BatchToken, queryName, CleanupMCode, mode, sheetName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        if (sheetName is not null) { _fixture.RegisterSheetForCleanup(sheetName); }
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, CleanupMCode,
            mode, sheetName, ["Val"], [[31], [83]]);

        if (delete)
        {
            RequireSuccess(_queries.Delete(_fixture.BatchToken, queryName));
            _fixture.ForgetPowerQuery(queryName);
            PowerQueryStateAssertions.AssertRemoved(_fixture, queryName);
        }
        else
        {
            RequireSuccess(_queries.Unload(_fixture.BatchToken, queryName));
            PowerQueryStateAssertions.AssertStored(_fixture, queryName, CleanupMCode,
                PowerQueryLoadMode.ConnectionOnly, null, ["Val"], [[31], [83]]);
        }

        Assert.DoesNotContain(RequireSuccess(_dataModel.ListTables(_fixture.BatchToken)).Tables,
            table => table.Name == queryName);
        var tables = _fixture.CreateCommands<ITableCommands>();
        Assert.DoesNotContain(RequireSuccess(tables.List(_fixture.BatchToken)).Tables,
            table => table.Name == queryName);
        if (sheetName is not null)
        {
            Assert.All(RequireSuccess(_commands.GetValues(
                _fixture.BatchToken, sheetName, "A1:A3")).Values,
                row => Assert.All(row, Assert.Null));
        }
        AssertInitialTable(guardName);
    }

    private static string UniqueCleanupName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];
}

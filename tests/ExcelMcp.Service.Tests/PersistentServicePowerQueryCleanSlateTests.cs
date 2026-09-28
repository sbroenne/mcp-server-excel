using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryCleanSlateTests(
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
    public void CreateThenDelete_LoadToTable_NoTraces()
    {
        var queryName = "PQ_CreateDelete_" + Guid.NewGuid().ToString("N")[..8];
        const string mCode =
            "let Source = #table({\"Val\"}, {{1}}) in Source";

        _queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToTable);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);
        _queries.Delete(_fixture.BatchToken, queryName);
        _fixture.ForgetPowerQuery(queryName);

        Assert.Empty(_queries.List(_fixture.BatchToken).Queries);
        Assert.Empty(_connections.List(_fixture.BatchToken).Connections);
        Assert.Empty(_tables.List(_fixture.BatchToken).Tables);
    }
}

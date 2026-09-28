using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal interface IPersistentChartCommands :
    IChartCommands,
    IChartConfigCommands;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceChartFormattingTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPersistentChartCommands _chartCommands;
    private readonly IPersistentTableCommands _tableCommands;
    private readonly string _sheetName;

    public PersistentServiceChartFormattingTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _chartCommands = ServiceCommandProxy.Create<IPersistentChartCommands>(fixture);
        _tableCommands = ServiceCommandProxy.Create<IPersistentTableCommands>(fixture);
        var batch = _fixture.BatchToken;
        _sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(
            batch,
            _sheetName,
            "A1:C6",
            [
                ["Category", "Series1", "Series2"],
                ["A", 10, 20],
                ["B", 15, 25],
                ["C", 20, 30],
                ["D", 25, 35],
                ["E", 30, 40],
            ]);
    }
}

using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal interface IPersistentTableCommands :
    ITableCommands,
    ITableColumnCommands,
    ISlicerCommands;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Tables")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceTablePreflightTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPersistentTableCommands _tableCommands;
    private readonly IRangeServiceCommands _rangeCommands;

    public PersistentServiceTablePreflightTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _tableCommands =
            ServiceCommandProxy.Create<IPersistentTableCommands>(fixture);
        _rangeCommands = ServiceCommandProxy.Create<IRangeServiceCommands>(fixture);

        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Sales");
        RequireSuccess(_rangeCommands.SetValues(
            batch,
            "Sales",
            "A1:D5",
            [["Region", "Product", "Amount", "Date"],
             ["North", "Widget", 100, "2025-01-15"],
             ["South", "Gadget", 250, "2025-02-20"],
             ["East", "Widget", 150, "2025-03-10"],
             ["West", "Gadget", 300, "2025-01-25"]]));
        RequireSuccess(_tableCommands.Create(
            batch, "Sales", "SalesTable", "A1:D5", true, "TableStyleMedium2"));
    }
}

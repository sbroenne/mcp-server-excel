using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal interface IPersistentPivotTableCommands :
    IPivotTableCommands,
    IPivotTableFieldCommands,
    IPivotTableCalcCommands,
    ISlicerCommands;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServicePivotTableTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPersistentPivotTableCommands _pivotCommands;
    private readonly IPersistentTableCommands _tableCommands;
    private readonly string _salesSheetName =
        $"Sales_{Guid.NewGuid():N}"[..31];

    public PersistentServicePivotTableTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _pivotCommands =
            ServiceCommandProxy.Create<IPersistentPivotTableCommands>(fixture);
        _tableCommands =
            ServiceCommandProxy.Create<IPersistentTableCommands>(fixture);
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, _salesSheetName);
        RequireSuccess(_commands.SetValues(
            batch,
            _salesSheetName,
            "A1:D6",
            [
                ["Region", "Product", "Sales", "Date"],
                ["North", "Widget", 100, "2025-01-15"],
                ["North", "Widget", 150, "2025-01-20"],
                ["South", "Gadget", 200, "2025-02-10"],
                ["North", "Gadget", 75, "2025-02-15"],
                ["South", "Widget", 125, "2025-03-05"],
            ]));
        RequireSuccess(_commands.SetNumberFormat(
            batch,
            _salesSheetName,
            "D2:D6",
            "m/d/yyyy"));
    }
}

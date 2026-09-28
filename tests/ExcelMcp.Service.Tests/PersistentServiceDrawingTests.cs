using Sbroenne.ExcelMcp.Core.Commands.Drawing;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Drawing")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceDrawingTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IDrawingCommands _drawingCommands;
    private readonly IRangeServiceCommands _rangeCommands;
    private readonly string _sheetName;

    public PersistentServiceDrawingTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _drawingCommands = ServiceCommandProxy.Create<IDrawingCommands>(fixture);
        _rangeCommands = ServiceCommandProxy.Create<IRangeServiceCommands>(fixture);
        _sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
    }
}

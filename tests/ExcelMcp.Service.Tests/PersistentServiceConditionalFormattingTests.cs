using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "ConditionalFormat")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceConditionalFormattingTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IConditionalFormattingCommands _conditionalFormattingCommands;
    private readonly IRangeCommands _rangeCommands;
    private readonly string _sheetName;

    public PersistentServiceConditionalFormattingTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _conditionalFormattingCommands =
            ServiceCommandProxy.Create<IConditionalFormattingCommands>(fixture);
        _sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _rangeCommands = ServiceCommandProxy.Create<IRangeCommands>(fixture);
        RequireSuccess(_rangeCommands.SetValues(_fixture.BatchToken, _sheetName, "A1:G41",
            Enumerable.Range(0, 41).Select(index => new List<object?>
                { index * 10 - 10, 1, 2, 3, 4, 5, index + 999 }).ToList()));
    }
}

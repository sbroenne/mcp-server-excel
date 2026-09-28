using Sbroenne.ExcelMcp.Core.Commands;
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

    public PersistentServiceConditionalFormattingTests(
        PersistentServiceWorkbookFixture fixture) :
        base(fixture)
    {
        _conditionalFormattingCommands =
            ServiceCommandProxy.Create<IConditionalFormattingCommands>(fixture);
        _fixture.CreateTestSheet(_fixture.BatchToken);
    }
}

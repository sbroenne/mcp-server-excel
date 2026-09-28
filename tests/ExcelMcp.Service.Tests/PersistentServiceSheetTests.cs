using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal interface IPersistentSheetCommands : ISheetCommands, ISheetStyleCommands;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Sheet")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceSheetTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPersistentSheetCommands _sheetCommands =
        ServiceCommandProxy.Create<IPersistentSheetCommands>(fixture);
}

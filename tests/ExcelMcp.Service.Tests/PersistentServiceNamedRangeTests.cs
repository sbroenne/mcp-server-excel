using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceNamedRangeTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly INamedRangeCommands _parameterCommands =
        ServiceCommandProxy.Create<INamedRangeCommands>(fixture);

    private static string CreateUniqueNamedRangeName() =>
        $"TestParam_{Guid.NewGuid():N}";
}

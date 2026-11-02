using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceNamedRangeTests(
    PersistentServiceWorkbookFixture fixture,
    ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly ITestOutputHelper _output = output;
    private readonly INamedRangeCommands _parameterCommands =
        ServiceCommandProxy.Create<INamedRangeCommands>(fixture);

    private static string CreateUniqueNamedRangeName() =>
        $"TestParam_{Guid.NewGuid():N}";
}

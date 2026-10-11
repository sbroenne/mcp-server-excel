using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceEmptyConnectionTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        var connections = _fixture.CreateCommands<IConnectionCommands>();
        var result = RequireSuccess(connections.List(_fixture.BatchToken));

        Assert.Empty(result.Connections);
        Assert.Equal(_fixture.WorkbookPath, result.FilePath);
    }
}

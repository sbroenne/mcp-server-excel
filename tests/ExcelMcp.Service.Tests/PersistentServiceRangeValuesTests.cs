using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public interface IRangeServiceCommands :
    IRangeCommands,
    IRangeEditCommands,
    IRangeFormatCommands,
    IRangeLinkCommands;

[System.Diagnostics.CodeAnalysis.SuppressMessage(
    "Design",
    "CA1051:Do not declare visible instance fields",
    Justification = "Protected fields preserve the original test source while routing through Service.")]
public abstract class PersistentServiceWorkbookTestBase(
    PersistentServiceWorkbookFixture fixture) : IAsyncLifetime
{
    protected readonly IRangeServiceCommands _commands =
        fixture.CreateCommands<IRangeServiceCommands>();
    protected readonly PersistentServiceWorkbookTestScope _fixture =
        fixture.CreateScope();

    public Task InitializeAsync() => Task.CompletedTask;

    protected static T RequireSuccess<T>(T result) where T : ResultBase
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage), result.ErrorMessage);
        return result;
    }

    public async Task DisposeAsync() => await _fixture.DisposeAsync();
}

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceRangeValuesTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>;

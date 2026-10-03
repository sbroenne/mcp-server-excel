using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public abstract class IsolatedServiceScreenshotTestBase :
    IAsyncLifetime,
    IAsyncDisposable
{
    private readonly PersistentServiceScreenshotFixture _owner = new();

    protected IScreenshotCommands _screenshotCommands { get; }
    protected ISheetStyleCommands _sheetCommands { get; }
    protected IRangeCommands _commands { get; }
    protected PersistentServiceWorkbookTestScope _fixture { get; }

    protected IsolatedServiceScreenshotTestBase()
    {
        _screenshotCommands = _owner.CreateCommands<IScreenshotCommands>();
        _sheetCommands = _owner.CreateCommands<ISheetStyleCommands>();
        _commands = _owner.CreateCommands<IRangeCommands>();
        _fixture = _owner.CreateScope();
    }

    public Task InitializeAsync() => _owner.InitializeAsync();

    protected static T RequireSuccess<T>(T result) where T : ResultBase
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage), result.ErrorMessage);
        return result;
    }

    public async Task DisposeAsync()
    {
        Exception? scopeFailure = null;
        try
        {
            await _fixture.DisposeAsync();
        }
        catch (Exception ex)
        {
            scopeFailure = ex;
        }

        try
        {
            await _owner.DisposeAsync();
        }
        catch (Exception ex) when (scopeFailure is not null)
        {
            throw new AggregateException(scopeFailure, ex);
        }

        if (scopeFailure is not null)
        {
            throw scopeFailure;
        }
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsync();
        GC.SuppressFinalize(this);
    }
}

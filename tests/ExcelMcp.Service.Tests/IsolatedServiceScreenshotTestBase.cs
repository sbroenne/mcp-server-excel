using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public abstract class IsolatedServiceScreenshotTestBase :
    IAsyncLifetime,
    IAsyncDisposable
{
    private readonly PersistentServiceScreenshotFixture _owner = new();

    protected IScreenshotCommands _screenshotCommands { get; }
    protected ISheetStyleCommands _sheetCommands { get; }
    protected PersistentServiceWorkbookTestScope _fixture { get; }

    protected IsolatedServiceScreenshotTestBase()
    {
        _screenshotCommands = _owner.CreateCommands<IScreenshotCommands>();
        _sheetCommands = _owner.CreateCommands<ISheetStyleCommands>();
        _fixture = _owner.CreateScope();
    }

    public Task InitializeAsync() => _owner.InitializeAsync();

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

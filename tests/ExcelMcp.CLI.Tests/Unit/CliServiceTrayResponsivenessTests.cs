using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "false")]
public sealed class CliServiceTrayResponsivenessTests
{
    [Fact(Timeout = 15000)]
    public async Task Tray_RefreshAndExitContinueWhileCloseIsBlocked()
    {
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var responsive = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var shutdown = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var finished = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var closeCalls = 0;
        var thread = new Thread(() =>
        {
            try
            {
                var refreshes = 0;
                using var form = new System.Windows.Forms.Form { ShowInTaskbar = false, Opacity = 0 };
                using var tray = new CliServiceTray(
                    () => { refreshes++; return []; },
                    (_, _) =>
                    {
                        Interlocked.Increment(ref closeCalls);
                        entered.TrySetResult();
                        release.Wait();
                        return Task.FromResult(new ServiceResponse { Success = true });
                    },
                    () =>
                    {
                        shutdown.TrySetResult();
                        return Task.FromResult<string?>(null);
                    });
                using var timer = new System.Windows.Forms.Timer { Interval = 50 };
                Task? close = null;
                timer.Tick += async (_, _) =>
                {
                    if (close is { IsCompleted: true })
                    {
                        timer.Stop();
                        await close;
                        form.Close();
                    }
                    else if (refreshes >= 2 && close is { IsCompleted: false })
                    {
                        timer.Stop();
                        responsive.TrySetResult();
                        await tray.ExitServiceAsync();
                        await close;
                        form.Close();
                    }
                };
                form.Shown += (_, _) =>
                {
                    timer.Start();
                    close = tray.CloseSessionAsync("test-session", false);
                };
                System.Windows.Forms.Application.Run(form);
                finished.TrySetResult();
            }
            catch (Exception ex)
            {
                finished.TrySetException(ex);
            }
        })
        { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        try
        {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
            await responsive.Task.WaitAsync(TimeSpan.FromSeconds(4));
            await shutdown.Task.WaitAsync(TimeSpan.FromSeconds(3));
            Assert.Equal(1, closeCalls);
        }
        finally
        {
            release.Set();
            await finished.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.True(thread.Join(TimeSpan.FromSeconds(1)));
        }
    }
}

using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.ComInterop.Session;
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
        var closeCalls = 0;
        var refreshes = 0;
        await CliServiceTrayTestHost.RunAsync(
            () => new CliServiceTray(
                () => { refreshes++; return []; },
                (_, _) =>
                {
                    Interlocked.Increment(ref closeCalls);
                    entered.TrySetResult();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(10)));
                    return Task.FromResult(new ServiceResponse { Success = true });
                },
                () =>
                {
                    shutdown.TrySetResult();
                    return Task.FromResult<string?>(null);
                }),
            async tray =>
            {
                using var timer = new System.Windows.Forms.Timer { Interval = 50 };
                timer.Tick += (_, _) =>
                {
                    if (refreshes >= 2)
                    {
                        timer.Stop();
                        responsive.TrySetResult();
                    }
                };
                timer.Start();
                var close = tray.CloseSessionAsync("test-session", false);
                Exception? failure = null;
                try
                {
                    await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
                    await responsive.Task.WaitAsync(TimeSpan.FromSeconds(4));
                    Assert.False(close.IsCompleted);
                    await tray.ExitServiceAsync();
                    await shutdown.Task.WaitAsync(TimeSpan.FromSeconds(3));
                    Assert.Equal(1, closeCalls);
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
                finally
                {
                    release.Set();
                    await CliServiceTrayTestHost.CaptureAsync(
                        () => close.WaitAsync(TimeSpan.FromSeconds(5)), ex => failure = Combine(failure, ex));
                }
                if (failure is not null)
                    ExceptionDispatchInfo.Capture(failure).Throw();
            });
    }

    [Fact(Timeout = 15000)]
    public async Task Tray_SaveOnExitKeepsUiResponsiveAndRejectsDuplicateExit()
    {
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var ticked = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var prompts = 0;
        var saves = 0;
        var shutdowns = 0;
        var notifications = new List<string>();
        await CliServiceTrayTestHost.RunAsync(
            () => new CliServiceTray(
                () => [new SessionDescriptor("owned", "owned.xlsx")],
                (sessionId, save) =>
                {
                    Assert.Equal("owned", sessionId);
                    Assert.True(save);
                    Interlocked.Increment(ref saves);
                    entered.TrySetResult();
                    Assert.True(release.Wait(TimeSpan.FromSeconds(10)));
                    return Task.FromResult(new ServiceResponse { Success = true });
                },
                () =>
                {
                    Interlocked.Increment(ref shutdowns);
                    return Task.FromResult<string?>(null);
                },
                (_, _, _, _) =>
                {
                    prompts++;
                    return DialogResult.Yes;
                },
                (_, message, _) => notifications.Add(message)),
            async tray =>
            {
                using var timer = new System.Windows.Forms.Timer { Interval = 50 };
                timer.Tick += (_, _) => ticked.TrySetResult();
                var exit = tray.ExitServiceAsync();
                Exception? failure = null;
                try
                {
                    await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
                    timer.Start();
                    await ticked.Task.WaitAsync(TimeSpan.FromSeconds(3));
                    Assert.False(exit.IsCompleted);
                    await tray.ExitServiceAsync();
                    Assert.Equal(1, prompts);
                    Assert.Equal(1, saves);
                    Assert.Equal(0, shutdowns);
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
                finally
                {
                    release.Set();
                    await CliServiceTrayTestHost.CaptureAsync(
                        () => exit.WaitAsync(TimeSpan.FromSeconds(5)), ex => failure = Combine(failure, ex));
                }
                if (failure is not null)
                    ExceptionDispatchInfo.Capture(failure).Throw();
                Assert.Equal(1, shutdowns);
                Assert.Empty(notifications);
            });
    }

    private static Exception Combine(Exception? primary, Exception cleanup) =>
        primary is null ? cleanup : new AggregateException(primary, cleanup);
}

internal static class CliServiceTrayTestHost
{
    internal static async Task RunAsync(Func<CliServiceTray> createTray, Func<CliServiceTray, Task> test)
    {
        var finished = new TaskCompletionSource<Exception?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var ready = new TaskCompletionSource<Form>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            Exception? threadFailure = null;
            Form? form = null;
            CliServiceTray? tray = null;
            try
            {
                form = new Form { ShowInTaskbar = false, Opacity = 0 };
                tray = createTray();
                var activeForm = form;
                var activeTray = tray;
                form.Shown += async (_, _) =>
                {
                    ready.TrySetResult(activeForm);
                    try
                    {
                        await test(activeTray);
                    }
                    catch (Exception ex)
                    {
                        threadFailure = ex;
                    }
                    finally
                    {
                        CaptureThreadCleanup(activeForm.Close);
                    }
                };
                Application.Run(form);
            }
            catch (Exception ex)
            {
                RecordThreadFailure(ex);
            }
            finally
            {
                CaptureThreadCleanup(() => tray?.Dispose());
                CaptureThreadCleanup(() => form?.Dispose());
                finished.TrySetResult(threadFailure);
            }

            void RecordThreadFailure(Exception ex) =>
                threadFailure = threadFailure is null ? ex : new AggregateException(threadFailure, ex);

            void CaptureThreadCleanup(Action action)
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    RecordThreadFailure(ex);
                }
            }
        })
        { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Exception? failure = null;
        try
        {
            failure = await finished.Task.WaitAsync(TimeSpan.FromMinutes(2));
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            if (!finished.Task.IsCompleted && ready.Task.IsCompletedSuccessfully)
            {
                await CaptureAsync(() =>
                {
                    ready.Task.Result.BeginInvoke((Action)ready.Task.Result.Close);
                    return Task.CompletedTask;
                }, Record);
            }
            await CaptureAsync(async () =>
            {
                var cleanupFailure = await finished.Task.WaitAsync(TimeSpan.FromSeconds(10));
                if (cleanupFailure is not null && !ReferenceEquals(cleanupFailure, failure))
                    Record(cleanupFailure);
            }, Record);
            if (!thread.Join(TimeSpan.FromSeconds(5)))
                Record(new TimeoutException("Tray STA thread did not exit."));
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();

        void Record(Exception ex) =>
            failure = failure is null ? ex : new AggregateException(failure, ex);
    }

    internal static async Task CaptureAsync(Func<Task> action, Action<Exception> record)
    {
        try
        {
            await action();
        }
        catch (Exception ex)
        {
            record(ex);
        }
    }
}

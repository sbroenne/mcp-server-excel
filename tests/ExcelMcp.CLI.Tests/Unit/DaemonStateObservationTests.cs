using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DaemonStateObservationTests
{
    [Fact]
    public async Task IsDaemonMutexHeld_LegacyDaemonMutexIsHeld_ReturnsTrue()
    {
        var pipeName = $"legacy-daemon-{Guid.NewGuid():N}";
        using var acquired = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var holder = Task.Run(() =>
        {
            using var legacyMutex = new Mutex(
                initiallyOwned: false,
                $"ExcelMcpCli_{pipeName}",
                out var createdNew);
            Assert.True(createdNew);
            legacyMutex.WaitOne();
            try
            {
                acquired.Set();
                release.Wait();
            }
            finally
            {
                legacyMutex.ReleaseMutex();
            }
        });

        Assert.True(acquired.Wait(TimeSpan.FromSeconds(5)));
        try
        {
            Assert.True(DaemonAutoStart.IsDaemonMutexHeld(pipeName));
        }
        finally
        {
            release.Set();
            await holder;
        }
    }

    [Fact]
    public async Task IsDaemonMutexHeld_LegacyDaemonMutexUsesCaseInsensitivePipeIdentity()
    {
        var legacyPipeName = $"legacy-daemon-case-{Guid.NewGuid():N}".ToLowerInvariant();
        var callerPipeName = legacyPipeName.ToUpperInvariant();
        using var acquired = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var holder = Task.Run(() =>
        {
            using var legacyMutex = new Mutex(
                initiallyOwned: false,
                $"ExcelMcpCli_{legacyPipeName}",
                out var createdNew);
            Assert.True(createdNew);
            legacyMutex.WaitOne();
            try
            {
                acquired.Set();
                release.Wait();
            }
            finally
            {
                legacyMutex.ReleaseMutex();
            }
        });

        Assert.True(acquired.Wait(TimeSpan.FromSeconds(5)));
        try
        {
            Assert.True(DaemonAutoStart.IsDaemonMutexHeld(callerPipeName));
        }
        finally
        {
            release.Set();
            await holder;
        }
    }

    [Fact]
    public void RecheckStartupAfterDaemonObservation_MarkerAppearsAfterInitialCheck_ReturnsStarting()
    {
        var observations = new Queue<(string Name, bool Value)>(
        [
            ("marker", false),
            ("daemon", true),
            ("marker", true)
        ]);

        var observation = DaemonConnectionPolicy.Observe(
            () => DequeueObservation(observations, "marker"),
            () => DequeueObservation(observations, "daemon"));

        Assert.True(observation.DaemonRunning);
        Assert.True(observation.StartupInProgress);
        Assert.Empty(observations);
    }

    [Fact]
    public void RecheckStartupAfterDaemonObservation_MarkerAppearsAfterStoppedObservation_ReturnsStarting()
    {
        var observations = new Queue<(string Name, bool Value)>(
        [
            ("marker", false),
            ("daemon", false),
            ("marker", true)
        ]);

        var observation = DaemonConnectionPolicy.Observe(
            () => DequeueObservation(observations, "marker"),
            () => DequeueObservation(observations, "daemon"));

        Assert.False(observation.DaemonRunning);
        Assert.True(observation.StartupInProgress);
        Assert.Empty(observations);
    }

    [Fact]
    public void Observe_InitialStartupMarkerDisappearsAfterDaemonObservation_UsesCompleteObservation()
    {
        var observations = new Queue<(string Name, bool Value)>(
        [
            ("marker", true),
            ("daemon", true),
            ("marker", false)
        ]);

        var observation = DaemonConnectionPolicy.Observe(
            () => DequeueObservation(observations, "marker"),
            () => DequeueObservation(observations, "daemon"));

        Assert.True(observation.DaemonRunning);
        Assert.True(observation.StartupInProgress);
        Assert.Empty(observations);
    }

    [Fact]
    public void Observe_InitialStartupMarkerPersistsAfterDaemonObservation_UsesCompleteObservation()
    {
        var observations = new Queue<(string Name, bool Value)>(
        [
            ("marker", true),
            ("daemon", false),
            ("marker", true)
        ]);

        var observation = DaemonConnectionPolicy.Observe(
            () => DequeueObservation(observations, "marker"),
            () => DequeueObservation(observations, "daemon"));

        Assert.False(observation.DaemonRunning);
        Assert.True(observation.StartupInProgress);
        Assert.Empty(observations);
    }

    [Fact]
    public void ShouldContinueStartupWait_DaemonReappearsAfterMutexGap_ReturnsTrue()
    {
        var daemonObservations = new Queue<bool>([false, true]);
        var startupObserved = true;

        Assert.False(daemonObservations.Dequeue());
        var daemonObservedAfterGap = daemonObservations.Dequeue();
        var closingMarkerProbeCount = 0;
        startupObserved = DaemonAutoStart.RecheckStartupAfterDaemonObservation(
            startupObserved,
            () =>
            {
                closingMarkerProbeCount++;
                return false;
            });
        var shouldWait = DaemonAutoStart.ShouldContinueStartupWait(
            startupObserved,
            daemonObservedAfterGap,
            startupDeadlineExpired: false);

        Assert.True(shouldWait);
        Assert.Equal(1, closingMarkerProbeCount);
        Assert.Empty(daemonObservations);
    }

    [Fact]
    public async Task EnsureAndConnectCoreAsync_FinalDaemonAndMarkerAppearAfterMutexGap_WaitsForStartup()
    {
        var daemonObservations = new Queue<bool>([true, false, true]);
        var markerObservations = new Queue<bool>([false, false, false, true]);
        var waitedForStartup = false;
        var runtime = new DaemonAutoStart.Runtime(
            PingAsync: (_, _) => Task.FromResult(false),
            IsDaemonMutexHeld: () => daemonObservations.Dequeue(),
            IsStartupInProgress: () => markerObservations.Dequeue(),
            TryStartDaemonAsync: (_, _) => throw new InvalidOperationException("The observer must not start another daemon."),
            WaitForResponsiveDaemonAsync: (deadline, _) =>
            {
                waitedForStartup = true;
                Assert.False(deadline.IsExpired);
                return Task.FromResult(true);
            });

        using var client = await DaemonAutoStart.EnsureAndConnectCoreAsync(
            $"race-test-{Guid.NewGuid():N}",
            OperationDeadline.Start(TimeSpan.FromSeconds(5)),
            runtime,
            CancellationToken.None);

        Assert.True(waitedForStartup);
        Assert.Empty(daemonObservations);
        Assert.Empty(markerObservations);
    }

    [Fact]
    public async Task EnsureAndConnectCoreAsync_SpawnedDaemonReadinessUsesOriginalDeadline()
    {
        TimeSpan? remainingBeforeLaunch = null;
        TimeSpan? remainingAfterLaunch = null;
        var runtime = new DaemonAutoStart.Runtime(
            PingAsync: async (_, cancellationToken) =>
            {
                await Task.Delay(TimeSpan.FromMilliseconds(200), cancellationToken);
                return false;
            },
            IsDaemonMutexHeld: () => false,
            IsStartupInProgress: () => throw new InvalidOperationException("No daemon was observed."),
            TryStartDaemonAsync: async (deadline, cancellationToken) =>
            {
                remainingBeforeLaunch = deadline.Remaining;
                await Task.Delay(TimeSpan.FromMilliseconds(250), cancellationToken);
                return DaemonAutoStart.StartOutcome.ObserveReadiness;
            },
            WaitForResponsiveDaemonAsync: (deadline, _) =>
            {
                remainingAfterLaunch = deadline.Remaining;
                return Task.FromResult(false);
            });

        await Assert.ThrowsAsync<TimeoutException>(() =>
            DaemonAutoStart.EnsureAndConnectCoreAsync(
                $"deadline-test-{Guid.NewGuid():N}",
                OperationDeadline.Start(TimeSpan.FromMilliseconds(700)),
                runtime,
                CancellationToken.None));

        Assert.NotNull(remainingBeforeLaunch);
        Assert.NotNull(remainingAfterLaunch);
        Assert.True(remainingAfterLaunch.Value < remainingBeforeLaunch.Value);
        Assert.InRange(
            remainingAfterLaunch.Value,
            TimeSpan.Zero,
            TimeSpan.FromMilliseconds(400));
    }

    [Fact]
    public async Task EnsureAndConnectCoreAsync_StartedDaemonAlreadyResponsive_DoesNotProbeAgain()
    {
        var runtime = new DaemonAutoStart.Runtime(
            PingAsync: (_, _) => Task.FromResult(false),
            IsDaemonMutexHeld: () => false,
            IsStartupInProgress: () => throw new InvalidOperationException("No daemon was observed."),
            TryStartDaemonAsync: (_, _) => Task.FromResult(DaemonAutoStart.StartOutcome.Ready),
            WaitForResponsiveDaemonAsync: (_, _) =>
                throw new InvalidOperationException("A ready daemon must not be probed again."));

        using var client = await DaemonAutoStart.EnsureAndConnectCoreAsync(
            $"ready-test-{Guid.NewGuid():N}",
            OperationDeadline.Start(TimeSpan.FromSeconds(1)),
            runtime,
            CancellationToken.None);
    }

    [Fact]
    public async Task StartDaemonProcess_ExpiredStartupDeadline_DoesNotLaunchProcess()
    {
        var deadline = OperationDeadline.Start(TimeSpan.FromMilliseconds(20));
        await Task.Delay(TimeSpan.FromMilliseconds(50));
        var launchAttempted = false;

        Assert.Throws<TimeoutException>(() =>
            DaemonAutoStart.StartDaemonProcess(
                deadline,
                CancellationToken.None,
                "test-daemon.exe",
                () =>
                {
                    launchAttempted = true;
                    return null;
                }));

        Assert.False(launchAttempted);
    }

    private static bool DequeueObservation(
        Queue<(string Name, bool Value)> observations,
        string expectedName)
    {
        var observation = observations.Dequeue();
        Assert.Equal(expectedName, observation.Name);
        return observation.Value;
    }
}

using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class DaemonConnectionPolicy
{
    internal const string RunningState = "running";
    internal const string StartingState = "starting";
    internal const string StoppedState = "stopped";
    internal const string UnresponsiveState = "unresponsive";

    internal static readonly TimeSpan InitialProbeTimeout = TimeSpan.FromSeconds(2);
    internal static readonly TimeSpan ControlTimeout = TimeSpan.FromSeconds(10);
    internal static readonly TimeSpan StartupReadyTimeout = TimeSpan.FromSeconds(30);

    internal static async Task<ServiceResponse> SendControlRequestAsync(
        string pipeName,
        ServiceRequest request,
        CancellationToken cancellationToken,
        TimeSpan? timeout = null)
    {
        var effectiveTimeout = timeout ?? ControlTimeout;
        using var client = new ServiceClient(
            pipeName,
            connectTimeout: effectiveTimeout,
            requestTimeout: effectiveTimeout);
        return await client.SendAsync(
            request,
            effectiveTimeout,
            cancellationToken);
    }

    internal static bool IsTransportFailure(ServiceResponse response) =>
        string.Equals(response.ErrorCategory, "Timeout", StringComparison.OrdinalIgnoreCase)
        || string.Equals(response.ErrorCategory, "ServiceUnavailable", StringComparison.OrdinalIgnoreCase);

    internal static DaemonObservation Observe(string pipeName)
    {
        return Observe(
            () => DaemonAutoStart.IsDaemonStartupInProgress(pipeName),
            () => DaemonAutoStart.IsDaemonMutexHeld(pipeName));
    }

    internal static DaemonObservation Observe(
        Func<bool> startupMarkerProbe,
        Func<bool> daemonProbe)
    {
        ArgumentNullException.ThrowIfNull(startupMarkerProbe);
        ArgumentNullException.ThrowIfNull(daemonProbe);

        var startupInProgress = startupMarkerProbe();
        var daemonRunning = daemonProbe();
        startupInProgress = DaemonAutoStart.RecheckStartupAfterDaemonObservation(
            startupInProgress,
            startupMarkerProbe);
        return new DaemonObservation(daemonRunning, startupInProgress);
    }

    internal static DaemonFailureState ResolveFailureState(
        string pipeName,
        ServiceResponse response)
    {
        if (!IsTransportFailure(response))
        {
            return new DaemonFailureState(RunningState, Running: true);
        }

        var observation = Observe(pipeName);
        if (observation.StartupInProgress)
        {
            return new DaemonFailureState(StartingState, Running: false);
        }

        if (observation.IsStopped)
        {
            return new DaemonFailureState(StoppedState, Running: false);
        }

        return new DaemonFailureState(UnresponsiveState, Running: true);
    }

    internal readonly record struct DaemonObservation(
        bool DaemonRunning,
        bool StartupInProgress)
    {
        internal bool IsStopped => !DaemonRunning && !StartupInProgress;
    }

    internal readonly record struct DaemonFailureState(string Name, bool Running);
}

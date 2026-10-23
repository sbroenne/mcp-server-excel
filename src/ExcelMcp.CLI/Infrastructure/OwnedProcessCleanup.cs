namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>Inspects recorded processes and removes only obsolete ownership information.</summary>
internal static class OwnedProcessCleanup
{
    internal sealed record ProcessSnapshot(
        DaemonProcessTracker.ProcessIdentity? DaemonProcess,
        bool DaemonMatched,
        IReadOnlyList<DaemonProcessTracker.ProcessIdentity> ExcelProcesses,
        DaemonProcessTracker.TrackingRecordStatus TrackingStatus,
        string? ErrorMessage)
    {
        internal IReadOnlyList<DaemonProcessTracker.ProcessIdentity> RetainedExcelProcesses { get; init; } = [];
    }

    internal sealed record CleanupResult(bool Success, bool DaemonMatched, string? ErrorMessage);
    internal sealed record RetainedExcelProcess(int ProcessId, DateTime StartTime, string WorkbookState = "unknown");

    internal const string RecoveryMessage =
        "Excel copies from an earlier CLI service remain open. Those older copies were not closed. " +
        "Their workbook contents and unsaved changes are unknown. " +
        "Use 'excelcli session discover', check complete and errors, then use " +
        "'excelcli session attach --workbook-id <ID>' for the intended workbook when available. " +
        "If it cannot be inspected, check the identified Excel copy before deciding what to do. " +
        "Ending Excel can lose unsaved work.";

    internal static ProcessSnapshot CaptureTrackedProcesses(string pipeName)
    {
        var read = DaemonProcessTracker.ReadProcessSnapshot(pipeName);
        if (read.Snapshot is not { } tracked)
            return new(null, false, [], read.Status, read.ErrorMessage);

        if (!DaemonProcessTracker.TryOpenMatchingProcess(tracked.DaemonProcess, out var daemon))
            return new(tracked.DaemonProcess, false, [], DaemonProcessTracker.TrackingRecordStatus.Unreadable,
                "The recorded CLI service process could not be checked.");
        var daemonMatched = daemon is not null;
        daemon?.Dispose();
        var current = new List<DaemonProcessTracker.ProcessIdentity>();
        var retained = new List<DaemonProcessTracker.ProcessIdentity>();
        foreach (var (source, destination) in new[]
        {
            (tracked.ExcelProcesses, current),
            (tracked.RetainedExcelProcesses, retained)
        })
        {
            foreach (var identity in source.Distinct())
            {
                if (!DaemonProcessTracker.TryOpenMatchingProcess(identity, out var process))
                    return new(tracked.DaemonProcess, daemonMatched, current,
                        DaemonProcessTracker.TrackingRecordStatus.Unreadable,
                        $"Recorded Excel process {identity.ProcessId} could not be checked.")
                    { RetainedExcelProcesses = retained };
                if (process is not null)
                {
                    process.Dispose();
                    destination.Add(identity);
                }
            }
        }
        return new(tracked.DaemonProcess, daemonMatched, current, read.Status, null)
        { RetainedExcelProcesses = retained };
    }

    internal static IReadOnlyList<RetainedExcelProcess> GetRetainedExcelProcesses(string pipeName, bool serviceRunning)
    {
        var snapshot = CaptureTrackedProcesses(pipeName);
        if (snapshot.TrackingStatus is DaemonProcessTracker.TrackingRecordStatus.Invalid
            or DaemonProcessTracker.TrackingRecordStatus.Unreadable)
            throw new InvalidOperationException(snapshot.ErrorMessage);
        return snapshot.RetainedExcelProcesses
            .Concat(serviceRunning ? [] : snapshot.ExcelProcesses)
            .Distinct()
            .Select(identity => new RetainedExcelProcess(
                identity.ProcessId, DateTime.FromFileTimeUtc(identity.StartedAtUtcFileTime)))
            .ToArray();
    }

    internal static Task<CleanupResult> CleanupAsync(string pipeName, CancellationToken cancellationToken) =>
        CleanupAsync(pipeName, CaptureTrackedProcesses(pipeName), cancellationToken);

    internal static Task<CleanupResult> CleanupAsync(
        string pipeName, ProcessSnapshot preShutdownSnapshot, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = CaptureTrackedProcesses(pipeName);
        if (snapshot.TrackingStatus is DaemonProcessTracker.TrackingRecordStatus.Invalid
            or DaemonProcessTracker.TrackingRecordStatus.Unreadable)
            return Task.FromResult(new CleanupResult(false, snapshot.DaemonMatched, snapshot.ErrorMessage));
        var success = snapshot.DaemonProcess != preShutdownSnapshot.DaemonProcess
            || snapshot.DaemonMatched
            || snapshot.ExcelProcesses.Count > 0
            || snapshot.RetainedExcelProcesses.Count > 0
            || snapshot.DaemonProcess is not { } daemon
            || DaemonProcessTracker.ClearIfDaemonMatches(pipeName, daemon);
        return Task.FromResult(new CleanupResult(success, snapshot.DaemonMatched,
            success ? null : "Obsolete CLI ownership information could not be removed. Ownership checks did not close Excel."));
    }
}

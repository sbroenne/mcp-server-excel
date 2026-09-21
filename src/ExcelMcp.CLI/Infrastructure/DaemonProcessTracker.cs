using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Tracks the daemon process for a specific pipe so the CLI can distinguish
/// "not running" from "running but unresponsive" during stop/recovery flows.
/// </summary>
internal static class DaemonProcessTracker
{
    private static readonly object TrackingFileLock = new();

    internal readonly record struct ProcessIdentity(int ProcessId, long StartedAtUtcFileTime);
    internal sealed record ProcessSnapshot(
        ProcessIdentity DaemonProcess,
        IReadOnlyList<ProcessIdentity> ExcelProcesses);
    internal enum TrackingRecordStatus
    {
        Missing,
        Available,
        Invalid,
        Unreadable
    }

    internal sealed record ProcessSnapshotReadResult(
        TrackingRecordStatus Status,
        ProcessSnapshot? Snapshot,
        string? ErrorMessage);

    private sealed class DaemonProcessRecord
    {
        public int ProcessId { get; set; }
        public long StartedAtUtcFileTime { get; set; }
        public List<TrackedProcessRecord> ExcelProcesses { get; set; } = [];
    }

    private sealed class TrackedProcessRecord
    {
        public int ProcessId { get; set; }
        public long StartedAtUtcFileTime { get; set; }
    }

    private sealed class TrackingMutexLease(Mutex mutex) : IDisposable
    {
        public void Dispose()
        {
            mutex.ReleaseMutex();
            mutex.Dispose();
        }
    }

    public static ProcessIdentity RegisterCurrentProcess(string pipeName)
    {
        var current = Process.GetCurrentProcess();
        return RegisterProcess(
            pipeName,
            current.Id,
            current.StartTime.ToUniversalTime().ToFileTimeUtc());
    }

    internal static ProcessIdentity RegisterProcess(
        string pipeName,
        int processId,
        long startedAtUtcFileTime)
    {
        var identity = new ProcessIdentity(processId, startedAtUtcFileTime);
        lock (TrackingFileLock)
        {
            using var trackingMutex = AcquireTrackingMutex(pipeName);
            Directory.CreateDirectory(GetTrackingDirectory());
            var record = new DaemonProcessRecord
            {
                ProcessId = processId,
                StartedAtUtcFileTime = startedAtUtcFileTime
            };

            WriteRecord(pipeName, record);
        }

        return identity;
    }

    /// <summary>
    /// Adds currently owned Excel identities to the selected daemon generation.
    /// Identities remain recorded until generation cleanup so a failed shutdown
    /// cannot lose a process that was untracked before it actually exited.
    /// </summary>
    internal static void RecordExcelProcesses(
        string pipeName,
        ProcessIdentity daemonIdentity,
        IReadOnlyCollection<ProcessIdentity> processIdentities)
    {
        lock (TrackingFileLock)
        {
            using var trackingMutex = AcquireTrackingMutex(pipeName);
            var readResult = ReadRecordCore(pipeName);
            if (readResult.Status != TrackingRecordStatus.Available
                || readResult.Record == null)
            {
                throw new InvalidOperationException(
                    readResult.ErrorMessage
                    ?? $"The daemon tracking record for pipe '{pipeName}' is unavailable.");
            }
            var record = readResult.Record;

            if (record.ProcessId != daemonIdentity.ProcessId
                || record.StartedAtUtcFileTime != daemonIdentity.StartedAtUtcFileTime)
            {
                throw new InvalidOperationException(
                    $"The daemon tracking record for pipe '{pipeName}' belongs to another process generation.");
            }

            var observedProcesses = processIdentities
                .Where(process => process.ProcessId > 0 && process.StartedAtUtcFileTime > 0)
                .Distinct()
                .Select(process => new TrackedProcessRecord
                {
                    ProcessId = process.ProcessId,
                    StartedAtUtcFileTime = process.StartedAtUtcFileTime
                })
                .ToList();
            record.ExcelProcesses = record.ExcelProcesses
                .Concat(observedProcesses)
                .DistinctBy(process => (process.ProcessId, process.StartedAtUtcFileTime))
                .ToList();
            WriteRecord(pipeName, record);
        }
    }

    internal static bool TryRecordExcelProcesses(
        string pipeName,
        ProcessIdentity daemonIdentity,
        IReadOnlyCollection<ProcessIdentity> processIdentities)
    {
        try
        {
            RecordExcelProcesses(pipeName, daemonIdentity, processIdentities);
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        catch (JsonException)
        {
            return false;
        }
        catch (Win32Exception)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    public static void Clear(string pipeName)
    {
        try
        {
            lock (TrackingFileLock)
            {
                using var trackingMutex = AcquireTrackingMutex(pipeName);
                var trackingFile = GetTrackingFilePath(pipeName);
                if (File.Exists(trackingFile))
                {
                    File.Delete(trackingFile);
                }
            }
        }
        catch
        {
            // Best-effort cleanup only.
        }
    }

    public static bool TryGetTrackedProcess(string pipeName, out Process? process)
    {
        process = null;
        if (!TryReadRecord(pipeName, out var record))
        {
            return false;
        }

        try
        {
            var candidate = Process.GetProcessById(record.ProcessId);
            if (candidate.HasExited)
            {
                candidate.Dispose();
                return false;
            }

            var startedAtUtcFileTime = candidate.StartTime.ToUniversalTime().ToFileTimeUtc();
            if (startedAtUtcFileTime != record.StartedAtUtcFileTime)
            {
                candidate.Dispose();
                return false;
            }

            process = candidate;
            return true;
        }
        catch
        {
            return false;
        }
    }

    internal static IReadOnlyList<ProcessIdentity> GetTrackedExcelProcessIdentities(string pipeName)
    {
        return TryGetProcessSnapshot(pipeName, out var snapshot)
            ? snapshot.ExcelProcesses
            : [];
    }

    internal static bool TryGetProcessSnapshot(string pipeName, out ProcessSnapshot snapshot)
    {
        var readResult = ReadProcessSnapshot(pipeName);
        if (readResult.Status != TrackingRecordStatus.Available
            || readResult.Snapshot == null)
        {
            snapshot = null!;
            return false;
        }

        snapshot = readResult.Snapshot;
        return true;
    }

    internal static ProcessSnapshotReadResult ReadProcessSnapshot(string pipeName)
    {
        try
        {
            lock (TrackingFileLock)
            {
                using var trackingMutex = AcquireTrackingMutex(pipeName);
                var readResult = ReadRecordCore(pipeName);
                if (readResult.Status != TrackingRecordStatus.Available
                    || readResult.Record == null)
                {
                    return new ProcessSnapshotReadResult(
                        readResult.Status,
                        null,
                        readResult.ErrorMessage);
                }

                return new ProcessSnapshotReadResult(
                    TrackingRecordStatus.Available,
                    new ProcessSnapshot(
                        new ProcessIdentity(
                            readResult.Record.ProcessId,
                            readResult.Record.StartedAtUtcFileTime),
                        readResult.Record.ExcelProcesses
                            .Select(tracked => new ProcessIdentity(
                                tracked.ProcessId,
                                tracked.StartedAtUtcFileTime))
                            .ToList()),
                    null);
            }
        }
        catch (IOException ex)
        {
            return UnreadableRecord(pipeName, ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            return UnreadableRecord(pipeName, ex);
        }
        catch (Win32Exception ex)
        {
            return UnreadableRecord(pipeName, ex);
        }
    }

    internal static bool ClearIfDaemonMatches(string pipeName, ProcessIdentity daemonIdentity)
    {
        try
        {
            lock (TrackingFileLock)
            {
                using var trackingMutex = AcquireTrackingMutex(pipeName);
                var readResult = ReadRecordCore(pipeName);
                if (readResult.Status == TrackingRecordStatus.Missing)
                {
                    return true;
                }

                if (readResult.Status != TrackingRecordStatus.Available
                    || readResult.Record == null)
                {
                    return false;
                }

                if (readResult.Record.ProcessId != daemonIdentity.ProcessId
                    || readResult.Record.StartedAtUtcFileTime != daemonIdentity.StartedAtUtcFileTime)
                {
                    return true;
                }

                var trackingFile = GetTrackingFilePath(pipeName);
                if (File.Exists(trackingFile))
                {
                    File.Delete(trackingFile);
                }

                return true;
            }
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static bool TryReadRecord(string pipeName, out DaemonProcessRecord record)
    {
        var readResult = ReadProcessSnapshot(pipeName);
        if (readResult.Status != TrackingRecordStatus.Available
            || readResult.Snapshot == null)
        {
            record = null!;
            return false;
        }

        record = new DaemonProcessRecord
        {
            ProcessId = readResult.Snapshot.DaemonProcess.ProcessId,
            StartedAtUtcFileTime = readResult.Snapshot.DaemonProcess.StartedAtUtcFileTime,
            ExcelProcesses = readResult.Snapshot.ExcelProcesses
                .Select(process => new TrackedProcessRecord
                {
                    ProcessId = process.ProcessId,
                    StartedAtUtcFileTime = process.StartedAtUtcFileTime
                })
                .ToList()
        };
        return true;
    }

    private static RecordReadResult ReadRecordCore(string pipeName)
    {
        var trackingFile = GetTrackingFilePath(pipeName);
        if (!File.Exists(trackingFile))
        {
            return new RecordReadResult(TrackingRecordStatus.Missing, null, null);
        }

        DaemonProcessRecord? parsed;
        try
        {
            var json = File.ReadAllText(trackingFile);
            parsed = JsonSerializer.Deserialize<DaemonProcessRecord>(
                json,
                DaemonTrackingJson.Options);
        }
        catch (JsonException ex)
        {
            return new RecordReadResult(
                TrackingRecordStatus.Invalid,
                null,
                $"The daemon tracking record '{trackingFile}' is malformed: {ex.Message}");
        }

        if (!IsValidRecord(parsed))
        {
            return new RecordReadResult(
                TrackingRecordStatus.Invalid,
                null,
                $"The daemon tracking record '{trackingFile}' contains invalid process ownership data.");
        }

        return new RecordReadResult(TrackingRecordStatus.Available, parsed, null);
    }

    private static bool IsValidRecord(DaemonProcessRecord? record) =>
        record is
        {
            ProcessId: > 0,
            StartedAtUtcFileTime: > 0,
            ExcelProcesses: not null
        }
        && record.ExcelProcesses.All(process =>
            process.ProcessId > 0
            && process.StartedAtUtcFileTime > 0);

    private static ProcessSnapshotReadResult UnreadableRecord(
        string pipeName,
        Exception exception) =>
        new(
            TrackingRecordStatus.Unreadable,
            null,
            $"The daemon tracking record for pipe '{pipeName}' could not be read: " +
            $"{exception.GetType().Name}: {exception.Message}");

    private sealed record RecordReadResult(
        TrackingRecordStatus Status,
        DaemonProcessRecord? Record,
        string? ErrorMessage);

    internal static bool TryOpenMatchingProcess(ProcessIdentity tracked, out Process? process)
    {
        process = null;
        try
        {
            var candidate = Process.GetProcessById(tracked.ProcessId);
            if (candidate.HasExited ||
                candidate.StartTime.ToUniversalTime().ToFileTimeUtc() != tracked.StartedAtUtcFileTime)
            {
                candidate.Dispose();
                return true;
            }

            process = candidate;
            return true;
        }
        catch (ArgumentException)
        {
            return true;
        }
        catch (InvalidOperationException)
        {
            return true;
        }
        catch (Win32Exception)
        {
            return false;
        }
        catch (NotSupportedException)
        {
            return false;
        }
    }

    private static void WriteRecord(string pipeName, DaemonProcessRecord record)
    {
        var trackingFile = GetTrackingFilePath(pipeName);
        var temporaryFile = $"{trackingFile}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllText(
                temporaryFile,
                JsonSerializer.Serialize(record, DaemonTrackingJson.Options));
            File.Move(temporaryFile, trackingFile, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryFile))
            {
                File.Delete(temporaryFile);
            }
        }
    }

    private static string GetTrackingDirectory() =>
        Path.Combine(Path.GetTempPath(), "ExcelMcp", "cli-daemon");

    private static TrackingMutexLease AcquireTrackingMutex(string pipeName)
    {
        var mutex = new Mutex(initiallyOwned: false, GetTrackingMutexName(pipeName));
        try
        {
            try
            {
                if (!mutex.WaitOne(TimeSpan.FromSeconds(5)))
                {
                    throw new IOException(
                        $"Timed out waiting for the process tracker lock for pipe '{pipeName}'.");
                }
            }
            catch (AbandonedMutexException)
            {
                // The abandoned mutex is acquired by the current thread.
            }

            return new TrackingMutexLease(mutex);
        }
        catch
        {
            mutex.Dispose();
            throw;
        }
    }

    internal static string GetTrackingFilePath(string pipeName)
    {
        return Path.Combine(
            GetTrackingDirectory(),
            $"{DaemonPipeIdentity.GetHash(pipeName)}.json");
    }

    internal static string GetTrackingMutexName(string pipeName) =>
        $"ExcelMcpCli_Tracker_{DaemonPipeIdentity.GetHash(pipeName)}";

}

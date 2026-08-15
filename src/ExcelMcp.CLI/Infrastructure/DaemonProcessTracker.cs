using System.ComponentModel;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Tracks the daemon process for a specific pipe so the CLI can distinguish
/// "not running" from "running but unresponsive" during stop/recovery flows.
/// </summary>
internal static class DaemonProcessTracker
{
    private static readonly object TrackingFileLock = new();

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

    public static void RegisterCurrentProcess(string pipeName)
    {
        var current = Process.GetCurrentProcess();
        RegisterProcess(pipeName, current.Id, current.StartTime.ToUniversalTime().ToFileTimeUtc());
    }

    internal static void RegisterProcess(string pipeName, int processId, long startedAtUtcFileTime)
    {
        lock (TrackingFileLock)
        {
            Directory.CreateDirectory(GetTrackingDirectory());
            var record = new DaemonProcessRecord
            {
                ProcessId = processId,
                StartedAtUtcFileTime = startedAtUtcFileTime
            };

            WriteRecord(pipeName, record);
        }
    }

    public static void UpdateExcelProcesses(string pipeName, IReadOnlyCollection<int> processIds)
    {
        UpdateExcelProcesses(pipeName, () => processIds);
    }

    internal static void UpdateExcelProcesses(
        string pipeName,
        Func<IReadOnlyCollection<int>> getProcessIds)
    {
        try
        {
            lock (TrackingFileLock)
            {
                if (!TryReadRecordCore(pipeName, out var record))
                {
                    return;
                }

                record.ExcelProcesses = getProcessIds()
                    .Distinct()
                    .Select(TryCreateProcessRecord)
                    .Where(process => process != null)
                    .Cast<TrackedProcessRecord>()
                    .ToList();
                WriteRecord(pipeName, record);
            }
        }
        catch (IOException)
        {
            // Tracking is best-effort and must never break session lifecycle.
        }
        catch (UnauthorizedAccessException)
        {
            // Tracking is best-effort and must never break session lifecycle.
        }
        catch (JsonException)
        {
            // A concurrent reader keeps the existing record for a later retry.
        }
        catch (Win32Exception)
        {
            // Process metadata became unavailable while creating the snapshot.
        }
    }

    public static void Clear(string pipeName)
    {
        try
        {
            lock (TrackingFileLock)
            {
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
                Clear(pipeName);
                return false;
            }

            var startedAtUtcFileTime = candidate.StartTime.ToUniversalTime().ToFileTimeUtc();
            if (startedAtUtcFileTime != record.StartedAtUtcFileTime)
            {
                candidate.Dispose();
                Clear(pipeName);
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

    public static IReadOnlyList<Process> GetTrackedExcelProcesses(string pipeName)
    {
        if (!TryReadRecord(pipeName, out var record))
        {
            return [];
        }

        var processes = new List<Process>();
        foreach (var tracked in record.ExcelProcesses)
        {
            var process = TryOpenMatchingProcess(tracked);
            if (process != null)
            {
                processes.Add(process);
            }
        }

        return processes;
    }

    private static bool TryReadRecord(string pipeName, out DaemonProcessRecord record)
    {
        record = null!;
        try
        {
            lock (TrackingFileLock)
            {
                return TryReadRecordCore(pipeName, out record);
            }
        }
        catch
        {
            return false;
        }
    }

    private static bool TryReadRecordCore(string pipeName, out DaemonProcessRecord record)
    {
        record = null!;
        var trackingFile = GetTrackingFilePath(pipeName);
        if (!File.Exists(trackingFile))
        {
            return false;
        }

        var json = File.ReadAllText(trackingFile);
        var parsed = JsonSerializer.Deserialize<DaemonProcessRecord>(json, ServiceProtocol.JsonOptions);
        if (parsed == null || parsed.ProcessId <= 0 || parsed.StartedAtUtcFileTime <= 0)
        {
            return false;
        }

        record = parsed;
        return true;
    }

    private static TrackedProcessRecord? TryCreateProcessRecord(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (process.HasExited)
            {
                return null;
            }

            return new TrackedProcessRecord
            {
                ProcessId = processId,
                StartedAtUtcFileTime = process.StartTime.ToUniversalTime().ToFileTimeUtc()
            };
        }
        catch (ArgumentException)
        {
            return null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }
    }

    private static Process? TryOpenMatchingProcess(TrackedProcessRecord tracked)
    {
        try
        {
            var process = Process.GetProcessById(tracked.ProcessId);
            if (process.HasExited ||
                process.StartTime.ToUniversalTime().ToFileTimeUtc() != tracked.StartedAtUtcFileTime)
            {
                process.Dispose();
                return null;
            }

            return process;
        }
        catch (ArgumentException)
        {
            return null;
        }
        catch (InvalidOperationException)
        {
            return null;
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
                JsonSerializer.Serialize(record, ServiceProtocol.JsonOptions));
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

    internal static string GetTrackingFilePath(string pipeName)
    {
        var nameBytes = Encoding.UTF8.GetBytes(pipeName);
        var safeName = Convert.ToHexString(SHA256.HashData(nameBytes));
        return Path.Combine(GetTrackingDirectory(), $"{safeName}.json");
    }
}

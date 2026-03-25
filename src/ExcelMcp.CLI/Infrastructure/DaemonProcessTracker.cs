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
    private sealed class DaemonProcessRecord
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
        Directory.CreateDirectory(GetTrackingDirectory());
        var record = new DaemonProcessRecord
        {
            ProcessId = processId,
            StartedAtUtcFileTime = startedAtUtcFileTime
        };

        File.WriteAllText(
            GetTrackingFilePath(pipeName),
            JsonSerializer.Serialize(record, ServiceProtocol.JsonOptions));
    }

    public static void Clear(string pipeName)
    {
        try
        {
            var trackingFile = GetTrackingFilePath(pipeName);
            if (File.Exists(trackingFile))
            {
                File.Delete(trackingFile);
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
            Clear(pipeName);
            return false;
        }
    }

    private static bool TryReadRecord(string pipeName, out DaemonProcessRecord record)
    {
        record = null!;

        try
        {
            var trackingFile = GetTrackingFilePath(pipeName);
            if (!File.Exists(trackingFile))
            {
                return false;
            }

            var json = File.ReadAllText(trackingFile);
            var parsed = JsonSerializer.Deserialize<DaemonProcessRecord>(json, ServiceProtocol.JsonOptions);
            if (parsed == null || parsed.ProcessId <= 0 || parsed.StartedAtUtcFileTime <= 0)
            {
                Clear(pipeName);
                return false;
            }

            record = parsed;
            return true;
        }
        catch
        {
            Clear(pipeName);
            return false;
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

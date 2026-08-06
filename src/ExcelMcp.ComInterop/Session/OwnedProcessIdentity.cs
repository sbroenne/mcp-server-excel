using System.Diagnostics;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Immutable identity for a process owned by the current automation instance.
/// A PID alone is not an identity because Windows can reuse it after exit.
/// </summary>
internal readonly record struct OwnedProcessIdentity(
    int ProcessId,
    long StartedAtUtcFileTime,
    string ProcessName,
    string ExecutablePath);

/// <summary>Read-only process facts used by the identity comparison seam.</summary>
internal readonly record struct ProcessIdentitySnapshot(
    int ProcessId,
    long StartedAtUtcFileTime,
    string ProcessName,
    string ExecutablePath);

/// <summary>
/// Opens and controls a process only while its PID, creation time, executable
/// name, and executable path still match the captured identity. Non-destructive
/// liveness checks use PID plus creation time and never grant control authority.
/// </summary>
internal static class OwnedProcessIdentityGuard
{
    public static bool TryCapture(int processId, out OwnedProcessIdentity identity)
    {
        identity = default;
        Process? process = null;
        try
        {
            process = Process.GetProcessById(processId);
            if (!TryReadSnapshot(process, out var snapshot))
            {
                return false;
            }

            identity = new OwnedProcessIdentity(
                snapshot.ProcessId,
                snapshot.StartedAtUtcFileTime,
                snapshot.ProcessName,
                snapshot.ExecutablePath);
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return false;
        }
        finally
        {
            process?.Dispose();
        }
    }

    public static bool IsExpectedExecutable(OwnedProcessIdentity identity, string processName) =>
        string.Equals(identity.ProcessName, processName, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(
            Path.GetFileNameWithoutExtension(identity.ExecutablePath),
            processName,
            StringComparison.OrdinalIgnoreCase);

    public static bool IsAlive(OwnedProcessIdentity identity)
    {
        Process? process = null;
        try
        {
            process = Process.GetProcessById(identity.ProcessId);
            return !process.HasExited &&
                   process.StartTime.ToUniversalTime().ToFileTimeUtc() == identity.StartedAtUtcFileTime;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return false;
        }
        finally
        {
            process?.Dispose();
        }
    }

    public static bool TryKill(OwnedProcessIdentity identity)
    {
        if (!TryOpenMatching(identity, out var process))
        {
            return false;
        }

        using (process)
        {
            try
            {
                process.Kill();
                return true;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                return false;
            }
        }
    }

    public static bool TryOpenMatching(OwnedProcessIdentity identity, out Process? process)
    {
        process = null;
        Process? candidate = null;
        try
        {
            candidate = Process.GetProcessById(identity.ProcessId);
            if (!TryReadSnapshot(candidate, out var snapshot) || !Matches(identity, snapshot))
            {
                return false;
            }

            process = candidate;
            candidate = null;
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return false;
        }
        finally
        {
            candidate?.Dispose();
        }
    }

    internal static bool Matches(OwnedProcessIdentity identity, ProcessIdentitySnapshot snapshot) =>
        identity.ProcessId == snapshot.ProcessId &&
        identity.StartedAtUtcFileTime == snapshot.StartedAtUtcFileTime &&
        string.Equals(identity.ProcessName, snapshot.ProcessName, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(identity.ExecutablePath, snapshot.ExecutablePath, StringComparison.OrdinalIgnoreCase);

    private static bool TryReadSnapshot(Process process, out ProcessIdentitySnapshot snapshot)
    {
        snapshot = default;
        if (process.HasExited)
        {
            return false;
        }

        string? executablePath = process.MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(executablePath))
        {
            return false;
        }

        snapshot = new ProcessIdentitySnapshot(
            process.Id,
            process.StartTime.ToUniversalTime().ToFileTimeUtc(),
            process.ProcessName,
            Path.GetFullPath(executablePath));
        return true;
    }
}

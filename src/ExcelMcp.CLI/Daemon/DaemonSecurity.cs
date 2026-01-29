using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;

namespace Sbroenne.ExcelMcp.CLI.Daemon;

/// <summary>
/// Security utilities for daemon named pipe communication.
/// Ensures per-user isolation via SID-based pipe names and ACLs.
/// </summary>
internal static class DaemonSecurity
{
    private static readonly string UserSid = WindowsIdentity.GetCurrent().User?.Value ?? "default";

    /// <summary>
    /// Gets the per-user pipe name.
    /// Format: excelcli-{USER_SID} to ensure isolation between users.
    /// </summary>
    public static string PipeName => $"excelcli-{UserSid}";

    /// <summary>
    /// Gets the per-user mutex name for single-instance enforcement.
    /// </summary>
    public static string MutexName => $"Global\\excelcli-daemon-{UserSid}";

    /// <summary>
    /// Gets the lock file path.
    /// </summary>
    public static string LockFilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "excelcli",
        "daemon.lock");

    /// <summary>
    /// Creates a secure named pipe server with ACLs restricting access to current user only.
    /// </summary>
    public static NamedPipeServerStream CreateSecureServer()
    {
        var pipeSecurity = new PipeSecurity();

        // Allow only the current user
        pipeSecurity.AddAccessRule(new PipeAccessRule(
            WindowsIdentity.GetCurrent().User!,
            PipeAccessRights.FullControl,
            AccessControlType.Allow));

        return NamedPipeServerStreamAcl.Create(
            PipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: NamedPipeServerStream.MaxAllowedServerInstances,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            inBufferSize: 4096,
            outBufferSize: 4096,
            pipeSecurity);
    }

    /// <summary>
    /// Creates a client connection to the daemon.
    /// </summary>
    public static NamedPipeClientStream CreateClient()
    {
        return new NamedPipeClientStream(
            ".",
            PipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous);
    }

    /// <summary>
    /// Tries to acquire the single-instance mutex.
    /// Returns the mutex if acquired, null if another instance exists.
    /// </summary>
    public static Mutex? TryAcquireSingleInstanceMutex()
    {
        var mutex = new Mutex(initiallyOwned: false, MutexName, out bool createdNew);

        if (!createdNew)
        {
            mutex.Dispose();
            return null;
        }

        try
        {
            if (!mutex.WaitOne(0))
            {
                mutex.Dispose();
                return null;
            }
            return mutex;
        }
        catch (AbandonedMutexException)
        {
            // Previous instance crashed, we can take over
            return mutex;
        }
    }

    /// <summary>
    /// Writes lock file with PID for status checking.
    /// </summary>
    public static void WriteLockFile(int pid)
    {
        var dir = Path.GetDirectoryName(LockFilePath)!;
        if (!Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }
        File.WriteAllText(LockFilePath, pid.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Reads daemon PID from lock file.
    /// </summary>
    public static int? ReadLockFilePid()
    {
        if (!File.Exists(LockFilePath))
        {
            return null;
        }

        try
        {
            var content = File.ReadAllText(LockFilePath).Trim();
            return int.TryParse(content, out var pid) ? pid : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Deletes the lock file.
    /// </summary>
    public static void DeleteLockFile()
    {
        try
        {
            if (File.Exists(LockFilePath))
            {
                File.Delete(LockFilePath);
            }
        }
        catch
        {
            // Best effort
        }
    }

    /// <summary>
    /// Checks if the daemon process is running based on lock file PID.
    /// </summary>
    public static bool IsDaemonProcessRunning()
    {
        var pid = ReadLockFilePid();
        if (!pid.HasValue)
        {
            return false;
        }

        try
        {
            var process = System.Diagnostics.Process.GetProcessById(pid.Value);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            // Process with this PID doesn't exist - clean up stale lock file
            DeleteLockFile();
            return false;
        }
        catch (InvalidOperationException)
        {
            // Process has exited
            DeleteLockFile();
            return false;
        }
        catch
        {
            // Other errors (e.g., access denied) - assume process might still be running
            // Don't delete lock file in case of transient errors
            return true;
        }
    }
}

using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Security utilities for ExcelMCP Service named pipe communication.
/// Ensures per-user isolation via SID-based pipe names and ACLs.
/// </summary>
/// <remarks>
/// <para><b>Security Model:</b></para>
/// <list type="bullet">
///   <item>User Isolation: Pipe name includes user SID - users cannot access each other's service instances</item>
///   <item>Windows ACLs: Named pipe restricts access to current user's SID via PipeSecurity</item>
///   <item>Local Only: Named pipes are local IPC - no network access possible</item>
/// </list>
/// <para><b>Not Enforced:</b></para>
/// <list type="bullet">
///   <item>Process Restriction: Any process running as the same user can connect to the service</item>
/// </list>
/// <para>
/// This is by design for a local automation tool. If malware runs under your user account,
/// it could already control Excel directly. The service does not elevate privileges.
/// See SECURITY.md for full documentation.
/// </para>
/// </remarks>
public static class ServiceSecurity
{
    private static readonly Lazy<string> LazyUserSid = new(() =>
    {
        try
        {
            return WindowsIdentity.GetCurrent().User?.Value ?? "default";
        }
        catch (Exception)
        {
            // WindowsIdentity may fail in containerized/restricted environments
            return "default";
        }
    });

    private static string UserSid => LazyUserSid.Value;

    /// <summary>
    /// Gets the per-user pipe name.
    /// Format: excelmcp-{USER_SID} to ensure isolation between users.
    /// </summary>
    public static string PipeName => $"excelmcp-{UserSid}";

    /// <summary>
    /// Gets the per-user mutex name for single-instance enforcement.
    /// </summary>
    public static string MutexName => $"Global\\ExcelMcpService-{UserSid}";

    /// <summary>
    /// Gets the lock file path.
    /// </summary>
    public static string LockFilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ExcelMCP",
        "service.lock");

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
    /// Creates a client connection to the service.
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

        bool acquired = false;
        try
        {
            acquired = mutex.WaitOne(0);
            if (!acquired)
            {
                return null;
            }
            return mutex;
        }
        catch (AbandonedMutexException)
        {
            // Previous instance crashed, we can take over
            acquired = true;
            return mutex;
        }
        finally
        {
            // Dispose mutex if we didn't acquire it (exception case)
            if (!acquired)
            {
                mutex.Dispose();
            }
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
    /// Reads service PID from lock file.
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
        catch (Exception)
        {
            // Lock file may be locked, corrupted, or inaccessible — treat as absent
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
        catch (Exception)
        {
            // Best-effort cleanup — lock file deletion is not critical
        }
    }

    /// <summary>
    /// Checks if the service process is running based on lock file PID.
    /// Guards against PID reuse by verifying the process name.
    /// </summary>
    public static bool IsServiceProcessRunning()
    {
        var pid = ReadLockFilePid();
        if (!pid.HasValue)
        {
            return false;
        }

        try
        {
            var process = System.Diagnostics.Process.GetProcessById(pid.Value);
            if (process.HasExited)
            {
                DeleteLockFile();
                return false;
            }

            // Guard against PID reuse: verify it's actually the service
            // Known service hosts: excelcli, Sbroenne.ExcelMcp.McpServer, dotnet (dev mode)
            var processName = process.ProcessName.ToLowerInvariant();
            if (processName != "excelcli" &&
                processName != "sbroenne.excelmcp.mcpserver" &&
                processName != "dotnet")
            {
                // Different process reused the PID - service is dead
                DeleteLockFile();
                return false;
            }

            return true;
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
        catch (Exception)
        {
            // Other errors (e.g., access denied) — assume process might still be running.
            // Don't delete lock file in case of transient errors.
            return true;
        }
    }
}



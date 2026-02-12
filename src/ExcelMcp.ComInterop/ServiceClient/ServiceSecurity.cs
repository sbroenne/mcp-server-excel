using System.IO.Pipes;
using System.Security.Principal;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Security utilities for ExcelMCP Service named pipe communication.
/// Ensures per-user isolation via SID-based pipe names.
/// This is the shared client-side portion used by both CLI and MCP Server.
/// </summary>
public static class ServiceSecurity
{
    private static readonly string UserSid = WindowsIdentity.GetCurrent().User?.Value ?? "default";

    /// <summary>
    /// Gets the per-user pipe name.
    /// Format: excelmcp-{USER_SID} to ensure isolation between users.
    /// </summary>
    public static string PipeName => $"excelmcp-{UserSid}";

    /// <summary>
    /// Gets the lock file path for checking service status.
    /// </summary>
    public static string LockFilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ExcelMCP",
        "service.lock");

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
                return false;
            }

            // Guard against PID reuse: verify it's actually the service
            // Process name will be "excelcli" (production) or "dotnet" (dev mode)
            var processName = process.ProcessName.ToLowerInvariant();
            return processName == "excelcli" || processName == "dotnet";
        }
        catch (ArgumentException)
        {
            // Process with this PID doesn't exist
            return false;
        }
        catch (Exception)
        {
            // Other errors (e.g., access denied) — assume service is not running
            return false;
        }
    }
}

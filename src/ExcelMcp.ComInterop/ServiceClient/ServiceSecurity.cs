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
}

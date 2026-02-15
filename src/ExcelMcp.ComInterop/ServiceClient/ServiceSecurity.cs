using System.IO.Pipes;
using System.Security.Principal;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Security utilities for ExcelMCP Service named pipe communication (client-side).
/// Provides pipe name generation and client connection creation.
/// </summary>
public static class ServiceSecurity
{
    private static readonly string UserSid = WindowsIdentity.GetCurrent().User?.Value ?? "default";

    /// <summary>
    /// Gets the pipe name for the MCP Server (per-process isolation).
    /// </summary>
    public static string GetMcpPipeName() => $"excelmcp-mcp-{UserSid}-{Environment.ProcessId}";

    /// <summary>
    /// Gets the pipe name for the CLI daemon (shared across CLI invocations for the same user).
    /// </summary>
    public static string GetCliPipeName() => $"excelmcp-cli-{UserSid}";

    /// <summary>
    /// Creates a client connection to a service pipe.
    /// </summary>
    public static NamedPipeClientStream CreateClient(string pipeName)
    {
        return new NamedPipeClientStream(
            ".",
            pipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous);
    }
}

using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Security utilities for ExcelMCP Service named pipe communication.
/// Ensures per-user isolation via SID-based pipe names and ACLs.
/// </summary>
/// <remarks>
/// <para><b>Pipe Name Strategy:</b></para>
/// <list type="bullet">
///   <item>MCP Server: excelmcp-mcp-{SID}-{PID} (per-process isolation, each instance independent)</item>
///   <item>CLI daemon: excelmcp-cli-{SID} (per-user, shared across CLI invocations)</item>
/// </list>
/// <para><b>Security Model:</b></para>
/// <list type="bullet">
///   <item>User Isolation: Pipe name includes user SID - users cannot access each other's service instances</item>
///   <item>Windows ACLs: Named pipe restricts access to current user's SID via PipeSecurity</item>
///   <item>Local Only: Named pipes are local IPC - no network access possible</item>
/// </list>
/// </remarks>
public static class ServiceSecurity
{
    private static readonly Lazy<string> LazyUserSid = new(() =>
    {
        var sid = WindowsIdentity.GetCurrent().User?.Value;
        if (string.IsNullOrEmpty(sid))
        {
            throw new InvalidOperationException(
                "Cannot determine current user SID. Named pipe security requires a valid SID for user isolation.");
        }
        return sid;
    });

    private static string UserSid => LazyUserSid.Value;

    /// <summary>
    /// Gets the pipe name for the MCP Server (per-process isolation).
    /// </summary>
    public static string GetMcpPipeName() => $"excelmcp-mcp-{UserSid}-{Environment.ProcessId}";

    /// <summary>
    /// Gets the pipe name for the CLI daemon (shared across CLI invocations for the same user).
    /// </summary>
    public static string GetCliPipeName() => $"excelmcp-cli-{UserSid}";

    /// <summary>
    /// Creates a secure named pipe server with ACLs restricting access to current user only.
    /// </summary>
    public static NamedPipeServerStream CreateSecureServer(string pipeName)
    {
        var pipeSecurity = new PipeSecurity();

        pipeSecurity.AddAccessRule(new PipeAccessRule(
            WindowsIdentity.GetCurrent().User!,
            PipeAccessRights.FullControl,
            AccessControlType.Allow));

        return NamedPipeServerStreamAcl.Create(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: NamedPipeServerStream.MaxAllowedServerInstances,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            inBufferSize: 4096,
            outBufferSize: 4096,
            pipeSecurity);
    }

    /// <summary>
    /// Creates a client connection to a service pipe.
    /// </summary>
    public static NamedPipeClientStream CreateClient(string pipeName)
    {
        return new NamedPipeClientStream(
            ".",
            pipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
    }
}

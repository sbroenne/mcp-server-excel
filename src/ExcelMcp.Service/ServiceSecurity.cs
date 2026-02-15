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
}

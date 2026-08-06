using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.ExcelMcp.Service.Workflow;

/// <summary>
/// Immutable identity supplied by the executable host that exposes workflow capabilities.
/// The host assembly is explicit so an in-process service cannot accidentally report a test
/// runner or a shared-library version as the installed runtime.
/// </summary>
public sealed class WorkflowRuntimeManifest
{
    /// <summary>Current contract version for the optimized workflow interface.</summary>
    public const string InterfaceVersion = "2";

    private WorkflowRuntimeManifest(
        string serverName,
        string serverVersion,
        string buildFingerprint,
        string toolProfile,
        string toolProfileVersion,
        string toolProfileFallback,
        IReadOnlyList<string> toolProfileTools,
        string toolProfileManifestHash)
    {
        ServerName = serverName;
        ServerVersion = serverVersion;
        BuildFingerprint = buildFingerprint;
        ToolProfile = toolProfile;
        ToolProfileVersion = toolProfileVersion;
        ToolProfileFallback = toolProfileFallback;
        ToolProfileTools = toolProfileTools;
        ToolProfileManifestHash = toolProfileManifestHash;
    }

    /// <summary>Name used by the host's protocol-level server identity.</summary>
    public string ServerName { get; }

    /// <summary>Version of the executable host, without an informational build suffix.</summary>
    public string ServerVersion { get; }

    /// <summary>Host build identifier from informational version metadata or module MVID.</summary>
    public string BuildFingerprint { get; }

    /// <summary>Identifier of the active tool profile.</summary>
    public string ToolProfile { get; }

    /// <summary>Version of the active tool-profile contract.</summary>
    public string ToolProfileVersion { get; }

    /// <summary>Profile clients can select when the active profile omits tools.</summary>
    public string ToolProfileFallback { get; }

    /// <summary>Ordinally sorted tool names exposed by the active profile.</summary>
    public IReadOnlyList<string> ToolProfileTools { get; }

    /// <summary>
    /// Lowercase SHA-256 hash of the active profile identifier, profile version, and tool list.
    /// It is always exactly 64 hexadecimal characters.
    /// </summary>
    public string ToolProfileManifestHash { get; }

    /// <summary>
    /// Creates a runtime manifest from a host-owned assembly. Callers must pass their own
    /// assembly rather than relying on <see cref="Assembly.GetEntryAssembly"/>, which is
    /// intentionally unreliable under test hosts and embedding scenarios.
    /// </summary>
    public static WorkflowRuntimeManifest Create(
        Assembly hostAssembly,
        string serverName,
        string toolProfile,
        IEnumerable<string> toolProfileTools,
        string toolProfileVersion = "1",
        string? toolProfileFallback = null)
    {
        ArgumentNullException.ThrowIfNull(hostAssembly);
        ArgumentException.ThrowIfNullOrWhiteSpace(serverName);
        ArgumentException.ThrowIfNullOrWhiteSpace(toolProfile);
        ArgumentNullException.ThrowIfNull(toolProfileTools);
        ArgumentException.ThrowIfNullOrWhiteSpace(toolProfileVersion);

        var tools = toolProfileTools
            .Select(tool => tool?.Trim() ?? throw new ArgumentException("Tool names cannot be null.", nameof(toolProfileTools)))
            .ToArray();
        if (tools.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("Tool names cannot be empty.", nameof(toolProfileTools));
        }

        var normalizedTools = tools
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)
            .ToArray();
        if (normalizedTools.Length != tools.Length)
        {
            throw new ArgumentException("Tool names must be unique.", nameof(toolProfileTools));
        }

        var informationalVersion = hostAssembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;
        var serverVersion = informationalVersion?.Split('+')[0]
            ?? hostAssembly.GetName().Version?.ToString()
            ?? "0.0.0";
        var buildFingerprint = GetBuildFingerprint(hostAssembly, informationalVersion);
        var fallback = string.IsNullOrWhiteSpace(toolProfileFallback)
            ? toolProfile
            : toolProfileFallback;

        return new WorkflowRuntimeManifest(
            serverName,
            serverVersion,
            buildFingerprint,
            toolProfile,
            toolProfileVersion,
            fallback,
            normalizedTools,
            CreateToolProfileManifestHash(toolProfile, toolProfileVersion, normalizedTools));
    }

    internal static WorkflowRuntimeManifest CreateServiceDefault() => Create(
        typeof(ExcelMcpService).Assembly,
        "excel-mcp-service",
        "service",
        Array.Empty<string>());

    private static string GetBuildFingerprint(Assembly hostAssembly, string? informationalVersion)
    {
        var separator = informationalVersion?.IndexOf('+', StringComparison.Ordinal) ?? -1;
        if (separator >= 0 && separator < informationalVersion!.Length - 1)
        {
            return informationalVersion[(separator + 1)..];
        }

        return hostAssembly.ManifestModule.ModuleVersionId.ToString("N");
    }

    private static string CreateToolProfileManifestHash(
        string toolProfile,
        string toolProfileVersion,
        IReadOnlyList<string> toolProfileTools)
    {
        var canonicalManifest = string.Join(
            '\n',
            ["workflow-tool-profile-v1", toolProfile, toolProfileVersion, .. toolProfileTools]);
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(canonicalManifest));
        return Convert.ToHexStringLower(hash);
    }
}

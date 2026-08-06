using System.Reflection;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Service.Workflow;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "Workflow")]
[Trait("Speed", "Fast")]
public sealed partial class WorkflowRuntimeManifestTests
{
    [Fact]
    public void Create_UsesTheExplicitHostAssemblyAndCanonicalProfileToolList()
    {
        var manifest = WorkflowRuntimeManifest.Create(
            typeof(Program).Assembly,
            "test-host",
            "test-profile",
            ["zeta", "alpha"],
            toolProfileVersion: "7",
            toolProfileFallback: "fallback");

        var reorderedManifest = WorkflowRuntimeManifest.Create(
            typeof(Program).Assembly,
            "another-host-name",
            "test-profile",
            ["alpha", "zeta"],
            toolProfileVersion: "7",
            toolProfileFallback: "another-fallback");

        Assert.Equal(ExpectedVersion(typeof(Program).Assembly), manifest.ServerVersion);
        Assert.Equal(ExpectedBuildFingerprint(typeof(Program).Assembly), manifest.BuildFingerprint);
        Assert.Equal(["alpha", "zeta"], manifest.ToolProfileTools);
        Assert.Equal(manifest.ToolProfileManifestHash, reorderedManifest.ToolProfileManifestHash);
        Assert.Matches(ProfileManifestHashPattern(), manifest.ToolProfileManifestHash);
    }

    [Fact]
    public void Create_ProfileHashChangesWhenTheActiveToolSurfaceChanges()
    {
        var baseManifest = WorkflowRuntimeManifest.Create(
            typeof(Program).Assembly,
            "test-host",
            "test-profile",
            ["alpha"]);
        var expandedManifest = WorkflowRuntimeManifest.Create(
            typeof(Program).Assembly,
            "test-host",
            "test-profile",
            ["alpha", "beta"]);

        Assert.NotEqual(baseManifest.ToolProfileManifestHash, expandedManifest.ToolProfileManifestHash);
    }

    private static string ExpectedVersion(Assembly assembly)
    {
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;
        return informationalVersion?.Split('+')[0]
            ?? assembly.GetName().Version?.ToString()
            ?? "0.0.0";
    }

    private static string ExpectedBuildFingerprint(Assembly assembly)
    {
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;
        var separator = informationalVersion?.IndexOf('+', StringComparison.Ordinal) ?? -1;
        return separator >= 0 && separator < informationalVersion!.Length - 1
            ? informationalVersion[(separator + 1)..]
            : assembly.ManifestModule.ModuleVersionId.ToString("N");
    }

    [GeneratedRegex("^[0-9a-f]{64}$", RegexOptions.CultureInvariant)]
    private static partial Regex ProfileManifestHashPattern();
}

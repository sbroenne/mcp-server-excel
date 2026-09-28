using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

public sealed class MacDistributionMetadataTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Fact]
    [Trait("Feature", "Distribution")]
    public void DistributionSurfaces_DeclareSeparateAppleSiliconAndIntelArtifacts()
    {
        var release = Read(".github/workflows/release.yml");
        var bootstrap = Read(".github/plugins/_shared/download.ps1.template");
        var extensionBuild = Read("vscode-extension/scripts/build-mcp-runtimes.mjs");
        var extensionRuntime = Read("vscode-extension/src/extension.ts");
        var extensionPackage = Read("vscode-extension/scripts/package-platforms.mjs");

        Assert.Contains("osx-arm64", release, StringComparison.Ordinal);
        Assert.Contains("runtime: osx-x64", release, StringComparison.Ordinal);
        Assert.Contains("@sbroenne/mcp-server-excel-darwin-x64", release, StringComparison.Ordinal);
        Assert.Contains("@sbroenne/excelcli-darwin-x64", release, StringComparison.Ordinal);
        Assert.Contains("ExcelMcp-MCP-Server-${{ env.VERSION }}-macos-x64.zip", release, StringComparison.Ordinal);
        Assert.Contains("ExcelMcp-CLI-${{ env.VERSION }}-macos-x64.zip", release, StringComparison.Ordinal);
        Assert.Contains("\"darwin-x64\"", release, StringComparison.Ordinal);
        Assert.Contains("slug: macos-x64", release, StringComparison.Ordinal);
        Assert.Contains("Architecture]::X64", bootstrap, StringComparison.Ordinal);
        Assert.Contains("Architecture]::Arm64", bootstrap, StringComparison.Ordinal);
        Assert.Contains("supports Windows x64 and macOS x64/Arm64 only", bootstrap, StringComparison.Ordinal);
        Assert.Contains("\"macos-arm64\"", bootstrap, StringComparison.Ordinal);
        Assert.Contains("\"macos-x64\"", bootstrap, StringComparison.Ordinal);
        Assert.Contains("runtime: 'osx-arm64'", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("runtime: 'osx-x64'", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("platform === 'darwin' && architecture === 'arm64'", extensionRuntime, StringComparison.Ordinal);
        Assert.Contains("platform === 'darwin' && architecture === 'x64'", extensionRuntime, StringComparison.Ordinal);
        Assert.Contains("Supported platforms are Windows x64 and macOS x64/Arm64", extensionRuntime, StringComparison.Ordinal);
        Assert.Contains("'darwin-arm64'", extensionPackage, StringComparison.Ordinal);
        Assert.Contains("'darwin-x64'", extensionPackage, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void ReleaseWorkflow_VerifiesIntelArchivesWithoutClaimingHardwareExecution()
    {
        var release = Read(".github/workflows/release.yml");
        var notarization = Read("scripts/Submit-MacNotarization.ps1");

        Assert.Contains("Test-DistributionPackages.ps1", release, StringComparison.Ordinal);
        Assert.Contains("Notarize Darwin VSIX payload", release, StringComparison.Ordinal);
        Assert.Contains("Intel macOS artifacts are build-verified; physical Intel Excel execution remains unverified", release, StringComparison.Ordinal);
        Assert.Contains("$configured = @(@(", notarization, StringComparison.Ordinal);
        Assert.Contains("submission.zip", notarization, StringComparison.Ordinal);
        Assert.Contains("\".zip\", \".pkg\", \".dmg\"", notarization, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void MacPackages_IncludeSignedArchitectureMatchedScreenCaptureHelper()
    {
        var release = Read(".github/workflows/release.yml");
        var helperBuild = Read("scripts/Build-MacScreenCaptureHelper.ps1");
        var npmBuild = Read("scripts/Build-NpmPackages.ps1");
        var extensionBuild = Read("vscode-extension/scripts/build-mcp-runtimes.mjs");
        var mcpbBuild = Read("mcpb/Build-McpBundle.ps1");

        Assert.Contains("-target \"$architecture-apple-macos14.0\"", helperBuild, StringComparison.Ordinal);
        Assert.Contains("$RuntimeIdentifier/helpers", helperBuild, StringComparison.Ordinal);
        Assert.Contains("helpers/excelmcp-screencapture", release, StringComparison.Ordinal);
        Assert.Contains("Sign-MacBinary.ps1", release, StringComparison.Ordinal);
        Assert.Contains("macOS runtime package requires the ScreenCaptureKit helper", npmBuild, StringComparison.Ordinal);
        Assert.Contains("Build-MacScreenCaptureHelper.ps1", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("Sign-MacBinary.ps1", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("Build-MacScreenCaptureHelper.ps1", mcpbBuild, StringComparison.Ordinal);
        Assert.Contains("server/helpers/excelmcp-screencapture", mcpbBuild, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void AgentSkillsBuild_UsesTheNativeCliForReferenceGeneration()
    {
        var buildScript = Read("scripts/Build-AgentSkills.ps1");

        Assert.Contains("if ($IsWindows) { \"excelcli.exe\" } else { \"excelcli\" }", buildScript, StringComparison.Ordinal);
        Assert.Contains("\"src/ExcelMcp.CLI/bin/Release/net10.0/$executableName\"", buildScript, StringComparison.Ordinal);
        Assert.DoesNotContain("net10.0-windows/excelcli.exe", buildScript, StringComparison.Ordinal);
    }

    private static string Read(string relativePath) =>
        File.ReadAllText(Path.Combine(RepoRoot, relativePath.Replace('/', Path.DirectorySeparatorChar)));

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate repository root.");
    }
}

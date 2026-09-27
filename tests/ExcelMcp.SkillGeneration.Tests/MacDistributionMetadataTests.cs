using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

public sealed class MacDistributionMetadataTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Fact]
    [Trait("Feature", "Distribution")]
    public void DistributionSurfaces_DeclareAppleSiliconAndFailClosedForIntel()
    {
        var release = Read(".github/workflows/release.yml");
        var bootstrap = Read(".github/plugins/_shared/download.ps1.template");
        var extensionBuild = Read("vscode-extension/scripts/build-mcp-runtimes.mjs");
        var extensionRuntime = Read("vscode-extension/src/extension.ts");
        var extensionPackage = Read("vscode-extension/scripts/package-platforms.mjs");

        Assert.Contains("osx-arm64", release, StringComparison.Ordinal);
        Assert.DoesNotContain("runtime: osx-x64", release, StringComparison.Ordinal);
        Assert.Contains("Architecture]::X64", bootstrap, StringComparison.Ordinal);
        Assert.Contains("Architecture]::Arm64", bootstrap, StringComparison.Ordinal);
        Assert.Contains("supports Windows x64 and Apple Silicon macOS only", bootstrap, StringComparison.Ordinal);
        Assert.Contains("\"macos-arm64\"", bootstrap, StringComparison.Ordinal);
        Assert.DoesNotContain("\"macos-x64\"", bootstrap, StringComparison.Ordinal);
        Assert.Contains("runtime: 'osx-arm64'", extensionBuild, StringComparison.Ordinal);
        Assert.DoesNotContain("runtime: 'osx-x64'", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("platform === 'darwin' && architecture === 'arm64'", extensionRuntime, StringComparison.Ordinal);
        Assert.Contains("Supported platforms are Windows x64 and Apple Silicon macOS", extensionRuntime, StringComparison.Ordinal);
        Assert.Contains("'darwin-arm64'", extensionPackage, StringComparison.Ordinal);
        Assert.DoesNotContain("'darwin-x64'", extensionPackage, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void ReleaseWorkflow_VerifiesArchivesAndDoesNotClaimIntelExecution()
    {
        var release = Read(".github/workflows/release.yml");
        var notarization = Read("scripts/Submit-MacNotarization.ps1");

        Assert.Contains("Test-DistributionPackages.ps1", release, StringComparison.Ordinal);
        Assert.Contains("Notarize Darwin VSIX payload", release, StringComparison.Ordinal);
        Assert.Contains("Darwin x64 is unsupported and no Intel macOS artifacts are published", release, StringComparison.Ordinal);
        Assert.Contains("$configured = @(@(", notarization, StringComparison.Ordinal);
        Assert.Contains("submission.zip", notarization, StringComparison.Ordinal);
        Assert.Contains("\".zip\", \".pkg\", \".dmg\"", notarization, StringComparison.Ordinal);
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

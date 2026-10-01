using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacDistributionMetadataTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Fact]
    [Trait("Feature", "Distribution")]
    public void DistributionSurfaces_DeclareAppleSiliconAndFailClosedForIntel()
    {
        var release = Read(".github/workflows/release.yml");
        var macPackages = Read("scripts/Build-MacReleasePackages.ps1");
        var launcher = Read("npm-packages/shared/launcher.js");
        var extensionBuild = Read("vscode-extension/scripts/build-mcp-runtimes.mjs");
        var extensionRuntime = Read("vscode-extension/src/extension.ts");
        var extensionPackage = Read("vscode-extension/scripts/package-platforms.mjs");

        Assert.Contains("Build-MacReleasePackages.ps1", release, StringComparison.Ordinal);
        Assert.Contains("osx-arm64", macPackages, StringComparison.Ordinal);
        Assert.DoesNotContain("osx-x64", macPackages, StringComparison.Ordinal);
        Assert.Contains("platform === 'win32' && (arch === 'x64' || arch === 'arm64')", launcher, StringComparison.Ordinal);
        Assert.Contains("platform === 'darwin' && arch === 'arm64'", launcher, StringComparison.Ordinal);
        Assert.Contains("supports Windows x64/Arm64 and Apple Silicon macOS", launcher, StringComparison.Ordinal);
        Assert.Contains("-darwin-arm64", launcher, StringComparison.Ordinal);
        Assert.DoesNotContain("-darwin-x64", launcher, StringComparison.Ordinal);
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
        var macPackages = Read("scripts/Build-MacReleasePackages.ps1");
        var notarization = Read("scripts/Submit-MacNotarization.ps1");

        Assert.Contains("Test-DistributionPackages.ps1", macPackages, StringComparison.Ordinal);
        Assert.Contains("excelmcp-$Version-darwin-arm64.vsix", macPackages, StringComparison.Ordinal);
        Assert.Contains("Apple Silicon", macPackages, StringComparison.Ordinal);
        Assert.Contains("$configured = @(@(", notarization, StringComparison.Ordinal);
        Assert.Contains("submission.zip", notarization, StringComparison.Ordinal);
        Assert.Contains("\".zip\", \".pkg\", \".dmg\"", notarization, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void MacPackages_IncludeSignedArchitectureMatchedScreenCaptureHelper()
    {
        var macPackages = Read("scripts/Build-MacReleasePackages.ps1");
        var helperBuild = Read("scripts/Build-MacScreenCaptureHelper.ps1");
        var npmBuild = Read("scripts/Build-NpmPackages.ps1");
        var extensionBuild = Read("vscode-extension/scripts/build-mcp-runtimes.mjs");
        var mcpbBuild = Read("mcpb/Build-McpBundle.ps1");

        Assert.Contains("-target \"$architecture-apple-macos14.0\"", helperBuild, StringComparison.Ordinal);
        Assert.Contains("$RuntimeIdentifier/helpers", helperBuild, StringComparison.Ordinal);
        Assert.DoesNotContain("osx-x64", helperBuild, StringComparison.Ordinal);
        Assert.Contains("helpers/excelmcp-screencapture", macPackages, StringComparison.Ordinal);
        Assert.Contains("Sign-MacBinary.ps1", macPackages, StringComparison.Ordinal);
        Assert.Contains("macOS runtime package requires the ScreenCaptureKit helper", npmBuild, StringComparison.Ordinal);
        Assert.Contains("Build-MacScreenCaptureHelper.ps1", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("Sign-MacBinary.ps1", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("Build-MacScreenCaptureHelper.ps1", mcpbBuild, StringComparison.Ordinal);
        Assert.Contains("server/helpers/excelmcp-screencapture", mcpbBuild, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void ReleaseWorkflow_RequiresMacNotarization()
    {
        var release = Read(".github/workflows/release.yml");
        var macPackages = Read("scripts/Build-MacReleasePackages.ps1");
        var notarization = Read("scripts/Submit-MacNotarization.ps1");
        var signing = Read("scripts/Initialize-MacSigning.ps1");
        var binarySigning = Read("scripts/Sign-MacBinary.ps1");
        var entitlements = Read("scripts/macos/ExcelMcp.Automation.entitlements.plist");

        Assert.Contains("Build-MacReleasePackages.ps1", release, StringComparison.Ordinal);
        Assert.Contains("Submit-MacNotarization.ps1", macPackages, StringComparison.Ordinal);
        Assert.Contains("Notarization credentials are required", notarization, StringComparison.Ordinal);
        Assert.Contains("AllowUnnotarized", notarization, StringComparison.Ordinal);
        Assert.Contains("Configured Developer ID signing identity", signing, StringComparison.Ordinal);
        Assert.Contains("SetUnixFileMode(", signing, StringComparison.Ordinal);
        Assert.Contains("$certificatePath,", signing, StringComparison.Ordinal);
        Assert.Contains("SetUnixFileMode(", notarization, StringComparison.Ordinal);
        Assert.Contains("[IO.Directory]::CreateDirectory(", notarization, StringComparison.Ordinal);
        Assert.Contains("[IO.UnixFileMode]::UserExecute", notarization, StringComparison.Ordinal);
        Assert.Contains("$keyPath,", notarization, StringComparison.Ordinal);
        Assert.Contains("AutomationClient", binarySigning, StringComparison.Ordinal);
        Assert.Contains("--entitlements", binarySigning, StringComparison.Ordinal);
        Assert.Contains("com.apple.security.automation.apple-events", entitlements, StringComparison.Ordinal);
        Assert.Contains("-AutomationClient", macPackages, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void ExtensionRuntimeBuild_SkipsDarwinHelperOnWindows()
    {
        var extensionBuild = Read("vscode-extension/scripts/build-mcp-runtimes.mjs");
        var extensionPackage = Read("vscode-extension/scripts/package-platforms.mjs");

        Assert.Contains("process.platform === 'win32'", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("process.platform !== 'darwin'", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("Windows hosts package only the Windows runtime", extensionBuild, StringComparison.Ordinal);
        Assert.Contains("process.platform === 'win32'", extensionPackage, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void McpbBuild_PreservesSiblingArtifactsAndDoesNotAdvertisePowerQuery()
    {
        var mcpbBuild = Read("mcpb/Build-McpBundle.ps1");

        Assert.DoesNotContain("Remove-Item -LiteralPath $OutputDir -Recurse", mcpbBuild, StringComparison.Ordinal);
        Assert.Contains("staging-$($Target.Slug)", mcpbBuild, StringComparison.Ordinal);
        Assert.DoesNotContain("plus Power Query list and view", mcpbBuild, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void AgentSkillsBuild_UsesTheNativeCliForReferenceGeneration()
    {
        var buildScript = Read("scripts/Build-AgentSkills.ps1");

        Assert.Contains("if ($IsWindows) { \"excelcli.exe\" } else { \"excelcli\" }", buildScript, StringComparison.Ordinal);
        Assert.Contains("if ($IsWindows) { \"net10.0-windows\" } else { \"net10.0\" }", buildScript, StringComparison.Ordinal);
        Assert.Contains("\"src/ExcelMcp.CLI/bin/Release/$targetFramework/$executableName\"", buildScript, StringComparison.Ordinal);
        Assert.Contains("$OutputDir = 'artifacts/generated-skills'", buildScript, StringComparison.Ordinal);
        Assert.Contains("$OutputDir = 'artifacts/skills'", buildScript, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Feature", "Distribution")]
    public void ExcelFreeRunner_SelectsPortableTestsCrossPlatform()
    {
        var runner = Read("scripts/Invoke-ExcelFreeTests.ps1");

        Assert.Contains("'Portable'", runner, StringComparison.Ordinal);
        Assert.Contains("@('Portable', 'SkillGeneration')", runner, StringComparison.Ordinal);
        Assert.Contains("RequiresExcel!=true&RunType!=OnDemand", runner, StringComparison.Ordinal);
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

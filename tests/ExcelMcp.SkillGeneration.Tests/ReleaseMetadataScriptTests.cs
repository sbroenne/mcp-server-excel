using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Integration tests for the release metadata updater used by the MCP Registry workflow.
/// </summary>
public sealed class ReleaseMetadataScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string UpdateMetadataScript = Path.Combine(
        RepoRoot,
        "scripts",
        "Update-McpRegistryMetadata.ps1");

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task UpdateMetadata_StampsSeparatedTopLevelAndPackageVersions()
    {
        var sandbox = CreateSandbox();
        try
        {
            var metadataPath = Path.Combine(sandbox, "server.json");
            File.Copy(
                Path.Combine(RepoRoot, "src", "ExcelMcp.McpServer", ".mcp", "server.json"),
                metadataPath);

            var result = await RunUpdaterAsync(metadataPath, "9.8.7");

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            using var document = JsonDocument.Parse(File.ReadAllText(metadataPath));
            var root = document.RootElement;
            Assert.Equal("9.8.7", root.GetProperty("version").GetString());

            var packages = root.GetProperty("packages").EnumerateArray().ToArray();
            var mcpPackage = Assert.Single(packages, package =>
                package.GetProperty("identifier").GetString() == "Sbroenne.ExcelMcp.McpServer");
            Assert.Equal("9.8.7", mcpPackage.GetProperty("version").GetString());
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "ReleaseMetadata")]
    public async Task UpdateMetadata_RejectsMissingMcpServerPackage()
    {
        var sandbox = CreateSandbox();
        try
        {
            var metadataPath = Path.Combine(sandbox, "server.json");
            await File.WriteAllTextAsync(metadataPath, """{"version":"1.0.0","packages":[]}""");

            var result = await RunUpdaterAsync(metadataPath, "9.8.7");

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("exactly one", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    private static string CreateSandbox()
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcpReleaseMetadata-{Guid.NewGuid():N}");
        Directory.CreateDirectory(sandbox);
        return sandbox;
    }

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

        throw new DirectoryNotFoundException("Could not locate repository root from test output directory.");
    }

    private static async Task<ScriptResult> RunUpdaterAsync(string metadataPath, string version)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = RepoRoot
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-ExecutionPolicy");
        startInfo.ArgumentList.Add("Bypass");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(UpdateMetadataScript);
        startInfo.ArgumentList.Add("-ServerJsonPath");
        startInfo.ArgumentList.Add(metadataPath);
        startInfo.ArgumentList.Add("-Version");
        startInfo.ArgumentList.Add(version);

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);

        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        await process.WaitForExitAsync(timeout.Token);

        return new ScriptResult(process.ExitCode, await stdout, await stderr);
    }

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}

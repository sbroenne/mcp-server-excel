using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Verifies that every Agent Skill shipped inside a built plugin carries a VERSION file stamped with
/// the version the plugin was built at.
/// </summary>
/// <remarks>
/// Regression guard: the published excel-cli plugin shipped without a VERSION file while excel-mcp
/// shipped with one. Two defects combined to cause it — the excel-cli Copy-AgentSkill call omitted
/// -Version, and Copy-AgentSkill only ever updated a VERSION file that already existed in the skill
/// source instead of creating one. These tests fail if either regresses, and they are agnostic to how
/// many skills a plugin ships so a future third skill is covered automatically.
/// </remarks>
public sealed class PluginSkillVersionTests
{
    private const string TestVersion = "9.9.9-skillversion";
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string BuildPluginsScript = Path.Combine(RepoRoot, "scripts", "Build-Plugins.ps1");
    private static readonly string BuildAgentSkillsScript = Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1");
    private static readonly string CopyVscodeSkillsScript = Path.Combine(RepoRoot, "scripts", "Copy-VscodeSkills.ps1");

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginSkillVersion")]
    public async Task BuildPlugins_StampsVersionFileIntoEverySkillDirectory()
    {
        var sandbox = CreateSandbox("plugin-skill-version");
        try
        {
            var outputDir = Path.Combine(sandbox, "built-plugins");

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                ["-Version", TestVersion, "-OutputDir", outputDir]);

            Assert.True(
                result.ExitCode == 0,
                $"Build-Plugins.ps1 failed with exit code {result.ExitCode}.{Environment.NewLine}{result.CombinedOutput}");

            var skillDirectories = Directory
                .GetDirectories(outputDir)
                .Select(pluginDir => Path.Combine(pluginDir, "skills"))
                .Where(Directory.Exists)
                .SelectMany(Directory.GetDirectories)
                .OrderBy(path => path, StringComparer.Ordinal)
                .ToList();

            // Guard against a vacuous pass if the output layout ever changes.
            Assert.True(
                skillDirectories.Count >= 2,
                $"Expected at least one skill per plugin, found {skillDirectories.Count} under {outputDir}.");

            foreach (var skillDirectory in skillDirectories)
            {
                var versionFile = Path.Combine(skillDirectory, "VERSION");

                Assert.True(
                    File.Exists(versionFile),
                    $"Built skill '{skillDirectory}' is missing a VERSION file.");

                Assert.Equal(TestVersion, File.ReadAllText(versionFile).Trim());
            }
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginSkillVersion")]
    public void CanonicalSkillSources_DoNotContainGeneratedVersionFiles()
    {
        var versionFiles = Directory
            .GetDirectories(Path.Combine(RepoRoot, "skills"), "excel-*")
            .Select(skillDirectory => Path.Combine(skillDirectory, "VERSION"));

        Assert.All(
            versionFiles,
            versionFile => Assert.False(
                File.Exists(versionFile),
                $"Canonical skill source contains generated package metadata: {versionFile}"));
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginSkillVersion")]
    public async Task BuildPlugins_RequiresExplicitVersion()
    {
        var sandbox = CreateSandbox("plugin-version-required");
        try
        {
            var result = await RunPowerShellFileAsync(
                BuildPluginsScript,
                ["-OutputDir", Path.Combine(sandbox, "built-plugins")]);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Version is required", result.CombinedOutput, StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginSkillVersion")]
    public async Task BuildAgentSkills_RequiresExplicitVersion()
    {
        var sandbox = CreateSandbox("agent-skills-version-required");
        try
        {
            var result = await RunPowerShellFileAsync(
                BuildAgentSkillsScript,
                ["-OutputDir", Path.GetRelativePath(RepoRoot, Path.Combine(sandbox, "skills"))]);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Version is required", result.CombinedOutput, StringComparison.Ordinal);
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginSkillVersion")]
    public async Task BuildAgentSkills_StampsVersionFileIntoEveryPackagedSkill()
    {
        var sandbox = CreateSandbox("agent-skills-version");
        try
        {
            var outputDir = Path.Combine(sandbox, "skills");
            var result = await RunPowerShellFileAsync(
                BuildAgentSkillsScript,
                [
                    "-Version",
                    TestVersion,
                    "-OutputDir",
                    Path.GetRelativePath(RepoRoot, outputDir)
                ]);

            Assert.True(
                result.ExitCode == 0,
                $"Build-AgentSkills.ps1 failed with exit code {result.ExitCode}.{Environment.NewLine}{result.CombinedOutput}");

            var zipPath = Path.Combine(outputDir, $"excel-skills-v{TestVersion}.zip");
            Assert.True(File.Exists(zipPath), $"Agent Skills ZIP was not created: {zipPath}");

            using var archive = ZipFile.OpenRead(zipPath);
            var versionEntries = archive.Entries
                .Where(entry => entry.FullName.EndsWith("/VERSION", StringComparison.Ordinal))
                .OrderBy(entry => entry.FullName, StringComparer.Ordinal)
                .ToList();

            Assert.Equal(
                ["skills/excel-cli/VERSION", "skills/excel-mcp/VERSION"],
                versionEntries.Select(entry => entry.FullName).ToArray());

            foreach (var entry in versionEntries)
            {
                using var reader = new StreamReader(entry.Open());
                Assert.Equal(TestVersion, (await reader.ReadToEndAsync()).Trim());
            }
        }
        finally
        {
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "PluginSkillVersion")]
    public async Task CopyVscodeSkills_CleansOutputAndStampsExtensionVersion()
    {
        var outputDir = Path.Combine(RepoRoot, "vscode-extension", "skills", "excel-mcp");
        try
        {
            DeleteDirectoryIfExists(outputDir);
            Directory.CreateDirectory(outputDir);
            File.WriteAllText(Path.Combine(outputDir, "stale.txt"), "stale");

            var result = await RunPowerShellFileAsync(CopyVscodeSkillsScript, []);

            Assert.True(
                result.ExitCode == 0,
                $"Copy-VscodeSkills.ps1 failed with exit code {result.ExitCode}.{Environment.NewLine}{result.CombinedOutput}");

            using var packageJson = JsonDocument.Parse(
                File.ReadAllText(Path.Combine(RepoRoot, "vscode-extension", "package.json")));
            var expectedVersion = packageJson.RootElement.GetProperty("version").GetString();

            Assert.False(File.Exists(Path.Combine(outputDir, "stale.txt")));
            Assert.True(File.Exists(Path.Combine(outputDir, "SKILL.md")));
            Assert.Equal(expectedVersion, File.ReadAllText(Path.Combine(outputDir, "VERSION")).Trim());
        }
        finally
        {
            DeleteDirectoryIfExists(outputDir);
        }
    }

    private static string CreateSandbox(string name)
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"excelmcp-{name}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private static void DeleteDirectoryIfExists(string path)
    {
        if (Directory.Exists(path))
        {
            Directory.Delete(path, recursive: true);
        }
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

    private static async Task<ProcessResult> RunPowerShellFileAsync(
        string scriptPath,
        IReadOnlyList<string> arguments,
        int timeoutMs = 120000)
    {
        var escapedScriptPath = scriptPath.Replace("'", "''");
        var escapedArguments = arguments
            .Select(argument => argument.Length > 0 && argument[0] == '-'
                ? argument
                : $"'{argument.Replace("'", "''")}'");
        var commandText = $"& '{escapedScriptPath}' {string.Join(" ", escapedArguments)}";

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
        startInfo.ArgumentList.Add("-Command");
        startInfo.ArgumentList.Add(commandText);

        using var process = new Process { StartInfo = startInfo };
        var stdout = new StringBuilder();
        var stderr = new StringBuilder();

        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data != null)
            {
                stdout.AppendLine(e.Data);
            }
        };

        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data != null)
            {
                stderr.AppendLine(e.Data);
            }
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        using var timeout = new CancellationTokenSource(timeoutMs);
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"PowerShell script '{scriptPath}' timed out after {timeoutMs}ms.");
        }

        return new ProcessResult(process.ExitCode, stdout.ToString(), stderr.ToString());
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}

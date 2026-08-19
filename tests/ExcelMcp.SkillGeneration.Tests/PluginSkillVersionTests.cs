using System.Diagnostics;
using System.Text;
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
    public void ExcelMcpSkillSource_KeepsVersionFile_BecauseBuildScriptsReadItAsFallbackVersion()
    {
        // Build-Plugins.ps1 and Build-AgentSkills.ps1 both fall back to skills/excel-mcp/VERSION when
        // no explicit -Version is supplied. Deleting it would silently break unversioned builds.
        var versionFile = Path.Combine(RepoRoot, "skills", "excel-mcp", "VERSION");

        Assert.True(File.Exists(versionFile), $"Canonical fallback version source is missing: {versionFile}");
        Assert.False(string.IsNullOrWhiteSpace(File.ReadAllText(versionFile)));
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

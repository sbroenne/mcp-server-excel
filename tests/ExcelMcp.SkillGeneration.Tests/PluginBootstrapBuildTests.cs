using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("GeneratedAssets")]
[Trait("RequiresExcel", "false")]
public sealed class PluginBootstrapBuildTests(ITestOutputHelper output)
{
    private const string AgentPluginSchema = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json";
    private const string AgentPluginMcpSchema = "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json";
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string BuildAgentSkillsScript = Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1");
    private static readonly string BuildPluginsScript = Path.Combine(RepoRoot, "scripts", "Build-Plugins.ps1");
    private static readonly string SyncPublishedRepoScript = Path.Combine(RepoRoot, "scripts", "Sync-PublishedPluginRepo.ps1");

    [Theory]
    [InlineData("excel-mcp")]
    [InlineData("excel-cli")]
    public void SourcePlugin_DoesNotShipGlobalInstaller(string pluginName)
    {
        var pluginRoot = Path.Combine(RepoRoot, ".github", "plugins", pluginName);
        Assert.Empty(Directory.GetFiles(pluginRoot, "install-global.ps1", SearchOption.AllDirectories));
    }

    [Theory]
    [InlineData("excel-mcp")]
    [InlineData("excel-cli")]
    [Trait("Feature", "SkillGeneration")]
    public void SourcePluginManifest_ConformsToAgentPluginsV1(string pluginName)
    {
        var pluginRoot = Path.Combine(RepoRoot, ".github", "plugins", pluginName);
        AssertAgentPluginManifest(pluginRoot, "0.0.0");
        AssertAgentSkill(Path.Combine(GeneratedAssetsFixture.SkillsDirectory, pluginName), pluginName);
    }

    [Theory]
    [InlineData("description: \"\"", "")]
    [InlineData("description: ''", "")]
    [InlineData("description: null", null)]
    [InlineData("description: \"\" # placeholder", "")]
    [InlineData("description: null # intentionally omitted", null)]
    [InlineData("description: \"Useful # skill\" # explanation", "Useful # skill")]
    [InlineData("description: >-\n  ", "")]
    [InlineData("description: >-\n  Useful skill", "Useful skill")]
    [Trait("Feature", "SkillGeneration")]
    public void SkillDescriptionMetadata_DecodesYamlScalar(string declaration, string? expected)
    {
        var lines = declaration.Split('\n');
        Assert.Equal(expected, DecodeYamlDescription(lines, 0));
    }

    [Fact]
    public void ExcelMcpSource_UsesPortableNpxConfiguration()
    {
        AssertPortableMcpConfiguration(Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp"));
    }

    [Fact]
    [Trait("Feature", "SkillGeneration")]
    public void PackagedReferences_AreReachableFromEachSkill()
    {
        foreach (var skill in new[] { "excel-cli", "excel-mcp" })
        {
            var root = Path.Combine(GeneratedAssetsFixture.SkillsDirectory, skill);
            var pending = new Stack<string>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            pending.Push(Path.Combine(root, "SKILL.md"));
            while (pending.TryPop(out var path))
            {
                path = Path.GetFullPath(path);
                if (!visited.Add(path))
                    continue;
                Assert.True(File.Exists(path), $"Missing linked file: {path}");
                foreach (Match match in Regex.Matches(File.ReadAllText(path), @"\]\(([^)]+)\)"))
                {
                    var target = match.Groups[1].Value.Split('#')[0];
                    if (target.Length > 0 && !target.Contains("://", StringComparison.Ordinal)
                        && target.EndsWith(".md", StringComparison.Ordinal))
                        pending.Push(Path.Combine(Path.GetDirectoryName(path)!, target));
                }
            }
            foreach (var reference in Directory.GetFiles(Path.Combine(root, "references"), "*.md", SearchOption.AllDirectories))
                Assert.Contains(Path.GetFullPath(reference), visited);
        }
    }

    [Fact]
    public async Task BuildPlugins_ProducesNpxOnlyPackages()
    {
        var sandbox = CreateSandbox("build");
        try
        {
            var outputDirectory = Path.Combine(sandbox, "built-plugins");
            const string version = "9.9.9-test";

            var result = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", version, "-OutputDir", outputDirectory]);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("[ok] excel-mcp - npx config and skill", result.Stdout, StringComparison.Ordinal);
            Assert.Contains("[ok] excel-cli - argument-safe npx wrapper and skill", result.Stdout, StringComparison.Ordinal);

            var mcpRoot = Path.Combine(outputDirectory, "excel-mcp");
            var cliRoot = Path.Combine(outputDirectory, "excel-cli");
            AssertAgentPluginManifest(mcpRoot, version);
            AssertAgentPluginManifest(cliRoot, version);
            AssertPortableMcpConfiguration(mcpRoot);
            Assert.True(File.Exists(Path.Combine(cliRoot, "bin", "start-cli.ps1")));
            Assert.False(File.Exists(Path.Combine(mcpRoot, "bin", "start-mcp.ps1")));
            Assert.False(File.Exists(Path.Combine(mcpRoot, "bin", "download.ps1")));
            Assert.False(File.Exists(Path.Combine(cliRoot, "bin", "download.ps1")));
            Assert.Empty(Directory.GetFiles(outputDirectory, "install-global.ps1", SearchOption.AllDirectories));

            AssertSkillDirectoryMatchesSource(
                Path.Combine(GeneratedAssetsFixture.SkillsDirectory, "excel-mcp"),
                Path.Combine(mcpRoot, "skills", "excel-mcp"), version);
            AssertSkillDirectoryMatchesSource(
                Path.Combine(GeneratedAssetsFixture.SkillsDirectory, "excel-cli"),
                Path.Combine(cliRoot, "skills", "excel-cli"), version);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task SyncPublishedPluginRepo_CopiesNpxOnlyPackages()
    {
        var sandbox = CreateSandbox("sync");
        try
        {
            var builtDirectory = Path.Combine(sandbox, "built");
            var publishedDirectory = Directory.CreateDirectory(Path.Combine(sandbox, "published")).FullName;
            File.WriteAllText(Path.Combine(publishedDirectory, "marketplace.json"), "{}");
            const string version = "9.9.10-test";

            var build = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", version, "-OutputDir", builtDirectory]);
            Assert.True(build.ExitCode == 0, build.CombinedOutput);
            var sync = await RunPowerShellFileAsync(
                SyncPublishedRepoScript,
                ["-PublishedRepoDir", publishedDirectory, "-BuiltPluginsDir", builtDirectory, "-Version", version]);

            Assert.True(sync.ExitCode == 0, sync.CombinedOutput);
            Assert.True(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-mcp", "mcp.json")));
            Assert.True(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-cli", "bin", "start-cli.ps1")));
            Assert.False(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-mcp", "bin", "download.ps1")));
            Assert.False(File.Exists(Path.Combine(publishedDirectory, "plugins", "excel-cli", "bin", "download.ps1")));
            Assert.Empty(Directory.GetFiles(publishedDirectory, "install-global.ps1", SearchOption.AllDirectories));

            var validation = await RunPowerShellFileAsync(
                Path.Combine(publishedDirectory, "tests", "Test-Plugins.ps1"), []);
            Assert.True(validation.ExitCode == 0, validation.CombinedOutput);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task SyncPublishedPluginRepo_RejectsMcpPackageWithoutConfiguration()
    {
        var sandbox = CreateSandbox("sync-missing-mcp-config");
        try
        {
            var builtDirectory = Path.Combine(sandbox, "built");
            var publishedDirectory = Directory.CreateDirectory(Path.Combine(sandbox, "published")).FullName;
            File.WriteAllText(Path.Combine(publishedDirectory, "marketplace.json"), "{}");

            var build = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", "9.9.11-test", "-OutputDir", builtDirectory]);
            Assert.True(build.ExitCode == 0, build.CombinedOutput);
            File.Delete(Path.Combine(builtDirectory, "excel-mcp", "mcp.json"));

            var sync = await RunPowerShellFileAsync(
                SyncPublishedRepoScript,
                ["-PublishedRepoDir", publishedDirectory, "-BuiltPluginsDir", builtDirectory, "-Version", "9.9.11-test"]);

            Assert.NotEqual(0, sync.ExitCode);
            Assert.Contains("excel-mcp is missing mcp.json", sync.Stderr, StringComparison.Ordinal);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Theory]
    [InlineData("excel-mcp", "bin")]
    [InlineData("excel-mcp", "com.github.copilot")]
    [InlineData("excel-cli", "bin")]
    [InlineData("excel-cli", "com.github.copilot")]
    public async Task SyncPublishedPluginRepo_RejectsRetiredGlobalInstaller(string pluginName, string directory)
    {
        var sandbox = CreateSandbox("sync-retired-installer");
        try
        {
            var builtDirectory = Path.Combine(sandbox, "built");
            var publishedDirectory = Directory.CreateDirectory(Path.Combine(sandbox, "published")).FullName;
            const string version = "9.9.12-test";

            var build = await RunPowerShellFileAsync(
                BuildPluginsScript, ["-Version", version, "-OutputDir", builtDirectory]);
            Assert.True(build.ExitCode == 0, build.CombinedOutput);
            var installerDirectory = Directory.CreateDirectory(Path.Combine(builtDirectory, pluginName, directory)).FullName;
            File.WriteAllText(Path.Combine(installerDirectory, "install-global.ps1"), "# Retired helper");

            var sync = await RunPowerShellFileAsync(
                SyncPublishedRepoScript,
                ["-PublishedRepoDir", publishedDirectory, "-BuiltPluginsDir", builtDirectory, "-Version", version]);

            Assert.NotEqual(0, sync.ExitCode);
            Assert.Contains("Global installation helpers are retired", sync.Stderr, StringComparison.Ordinal);
            Assert.Empty(Directory.GetFileSystemEntries(publishedDirectory));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    [SupportedOSPlatform("windows")]
    public async Task StartCliWrapper_EscapesArgumentsSoTheyRoundTripThroughWin32Parsing()
    {
        var sandbox = CreateSandbox("argument-fidelity");
        try
        {
            var harness = Path.Combine(sandbox, "harness.ps1");
            File.WriteAllText(harness, """
                param([Parameter(Mandatory = $true)][string]$ScriptPath)
                $tokens = $null
                $errors = $null
                $ast = [Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
                if ($errors.Count) { throw "Parse errors in $ScriptPath" }
                $definition = $ast.Find({
                    param($node)
                    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq 'ConvertTo-NativeArgument'
                }, $true)
                if ($null -eq $definition) { throw "ConvertTo-NativeArgument was not found" }
                . ([scriptblock]::Create($definition.Extent.Text))
                $cases = @('[["Name","Amount"]]', 'has space', 'quote"inside', '', '{"value":"a b"}')
                Write-Output (($cases | ForEach-Object { ConvertTo-NativeArgument -Value $_ }) -join ' ')
                """);

            var result = await RunPowerShellFileAsync(
                harness, ["-ScriptPath", Path.Combine(RepoRoot, ".github", "plugins", "excel-cli", "bin", "start-cli.ps1")]);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            var parsed = SplitCommandLine($"excelcli.exe {result.Stdout.Trim()}").Skip(1).ToArray();
            Assert.Equal(["[[\"Name\",\"Amount\"]]", "has space", "quote\"inside", "", "{\"value\":\"a b\"}"], parsed);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task StartCliWrapper_UsesNpxAndPreservesJsonArgument()
    {
        var sandbox = CreateSandbox("npx-wrapper");
        try
        {
            var wrapper = Path.Combine(RepoRoot, ".github", "plugins", "excel-cli", "bin", "start-cli.ps1");
            File.WriteAllText(Path.Combine(sandbox, "npx.cmd"), "@echo off\r\n");
            var npmBin = Directory.CreateDirectory(Path.Combine(sandbox, "node_modules", "npm", "bin")).FullName;
            File.WriteAllText(Path.Combine(npmBin, "npx-cli.js"), """
                if (process.argv[2] !== "-y" || process.argv[3] !== "@sbroenne/excelcli@latest") {
                    throw new Error("Unexpected npx arguments");
                }
                process.stdout.write(process.argv[4]);
                """);
            const string json = """[["Name","Amount"],["Widget",1500]]""";
            var harness = Path.Combine(sandbox, "invoke.ps1");
            File.WriteAllText(harness, """
                param([string]$Wrapper, [string]$NpxDirectory, [string]$Json)
                $env:PATH = "$NpxDirectory;$env:PATH"
                & $Wrapper $Json
                exit $LASTEXITCODE
                """);

            var result = await RunPowerShellFileAsync(
                harness, ["-Wrapper", wrapper, "-NpxDirectory", sandbox, "-Json", json]);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Equal(json, result.Stdout.Trim());
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [SupportedOSPlatform("windows")]
    private static string[] SplitCommandLine(string commandLine)
    {
        var argv = CommandLineToArgvW(commandLine, out var count);
        if (argv == IntPtr.Zero) { throw new Win32Exception(Marshal.GetLastWin32Error()); }
        try
        {
            return Enumerable.Range(0, count)
                .Select(index => Marshal.PtrToStringUni(Marshal.ReadIntPtr(argv, index * IntPtr.Size)) ?? "")
                .ToArray();
        }
        finally { LocalFree(argv); }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CommandLineToArgvW(string lpCmdLine, out int pNumArgs);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LocalFree(IntPtr hMem);

    private static void AssertAgentPluginManifest(string pluginRoot, string expectedVersion)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.Combine(pluginRoot, "plugin.json")));
        var root = document.RootElement;
        Assert.Equal(AgentPluginSchema, root.GetProperty("$schema").GetString());
        Assert.Equal(Path.GetFileName(pluginRoot), root.GetProperty("name").GetString());
        Assert.Equal(expectedVersion, root.GetProperty("version").GetString());
    }

    private static void AssertPortableMcpConfiguration(string pluginRoot)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.Combine(pluginRoot, "mcp.json")));
        var root = document.RootElement;
        Assert.Equal(AgentPluginMcpSchema, root.GetProperty("$schema").GetString());
        var server = root.GetProperty("mcpServers").GetProperty("excel-mcp");
        Assert.Equal("stdio", server.GetProperty("type").GetString());
        Assert.Equal("npx", server.GetProperty("command").GetString());
        Assert.Equal(["-y", "@sbroenne/mcp-server-excel@latest"],
            server.GetProperty("args").EnumerateArray().Select(value => value.GetString()!).ToArray());
    }

    private static void AssertAgentSkill(string skillRoot, string expectedName)
    {
        var content = File.ReadAllText(Path.Combine(skillRoot, "SKILL.md"));
        var header = Regex.Match(content, @"\A---\r?\n(?<header>.*?)\r?\n---(?:\r?\n|\z)", RegexOptions.Singleline);
        Assert.True(header.Success, "Skill metadata header is missing.");
        var metadata = header.Groups["header"].Value;
        Assert.Matches($@"(?m)^name:\s*{Regex.Escape(expectedName)}\s*$", metadata);
        var lines = metadata.Split('\n').Select(line => line.TrimEnd('\r')).ToArray();
        var index = Array.FindIndex(lines, line => line.StartsWith("description:", StringComparison.Ordinal));
        Assert.True(index >= 0, "Skill description metadata is missing.");
        var description = DecodeYamlDescription(lines, index);
        Assert.False(string.IsNullOrWhiteSpace(description), "Skill description metadata is empty.");
    }

    private static string? DecodeYamlDescription(string[] lines, int index)
    {
        var value = StripYamlInlineComment(lines[index]["description:".Length..].Trim());
        if (value is ">" or ">-" or ">+" or "|" or "|-" or "|+")
            return string.Join(" ", lines.Skip(index + 1)
                .TakeWhile(line => string.IsNullOrWhiteSpace(line) || char.IsWhiteSpace(line[0]))
                .Select(line => line.Trim()))
                .Trim();
        if (value is "~" || value.Equals("null", StringComparison.OrdinalIgnoreCase))
            return null;
        if (value.Length >= 2 && value[0] == '\'' && value[^1] == '\'')
            return value[1..^1].Replace("''", "'", StringComparison.Ordinal);
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
            return JsonSerializer.Deserialize<string>(value);
        return value;
    }

    private static string StripYamlInlineComment(string value)
    {
        var inSingleQuotes = false;
        var inDoubleQuotes = false;
        var escaped = false;
        for (var i = 0; i < value.Length; i++)
        {
            var character = value[i];
            if (inDoubleQuotes && character == '\\' && !escaped)
            {
                escaped = true;
                continue;
            }
            if (character == '"' && !inSingleQuotes && !escaped)
                inDoubleQuotes = !inDoubleQuotes;
            else if (character == '\'' && !inDoubleQuotes)
            {
                if (inSingleQuotes && i + 1 < value.Length && value[i + 1] == '\'')
                    i++;
                else
                    inSingleQuotes = !inSingleQuotes;
            }
            else if (character == '#' && !inSingleQuotes && !inDoubleQuotes
                     && (i == 0 || char.IsWhiteSpace(value[i - 1])))
                return value[..i].TrimEnd();
            escaped = false;
        }
        return value;
    }

    private static void AssertSkillDirectoryMatchesSource(
        string sourceRoot, string builtRoot, string expectedVersion)
    {
        var sourceFiles = Directory.GetFiles(sourceRoot, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(sourceRoot, path))
            .Where(path => path != "VERSION")
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        var builtFiles = Directory.GetFiles(builtRoot, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(builtRoot, path))
            .Where(path => path != "VERSION")
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(sourceFiles, builtFiles);
        foreach (var path in sourceFiles)
        {
            Assert.Equal(
                File.ReadAllBytes(Path.Combine(sourceRoot, path)),
                File.ReadAllBytes(Path.Combine(builtRoot, path)));
        }
        Assert.Equal(expectedVersion, File.ReadAllText(Path.Combine(builtRoot, "VERSION")).Trim());
    }

    private static string CreateSandbox(string name)
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcp-{name}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(sandbox);
        return sandbox;
    }

    private static void DeleteDirectoryIfExists(string path)
    {
        if (Directory.Exists(path)) { Directory.Delete(path, true); }
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate repository root.");
    }

    private async Task<ProcessResult> RunPowerShellFileAsync(
        string scriptPath,
        IReadOnlyList<string> arguments,
        Dictionary<string, string>? environmentVariables = null,
        int timeoutMs = 30000)
    {
        if (scriptPath == BuildPluginsScript || scriptPath == BuildAgentSkillsScript)
        {
            arguments = [.. arguments, "-SkillsDirectory", GeneratedAssetsFixture.SkillsDirectory];
        }
        var escapedArguments = arguments.Select(argument =>
            argument.Length > 0 && argument[0] == '-' ? argument : $"'{argument.Replace("'", "''")}'");
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
        startInfo.ArgumentList.Add($"& '{scriptPath.Replace("'", "''")}' {string.Join(" ", escapedArguments)}");
        foreach (var ambientName in new[] { "COPILOT_AGENT_SESSION_ID", "PLUGIN_DATA", "GITHUB_TOKEN", "GH_TOKEN" })
        {
            startInfo.Environment.Remove(ambientName);
        }
        if (environmentVariables != null)
        {
            foreach (var (key, value) in environmentVariables) { startInfo.Environment[key] = value; }
        }

        using var process = new Process { StartInfo = startInfo };
        var stdout = new StringBuilder();
        var stderr = new StringBuilder();
        process.OutputDataReceived += (_, eventArgs) => { if (eventArgs.Data != null) { stdout.AppendLine(eventArgs.Data); } };
        process.ErrorDataReceived += (_, eventArgs) => { if (eventArgs.Data != null) { stderr.AppendLine(eventArgs.Data); } };
        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        using var timeout = new CancellationTokenSource(timeoutMs);
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            process.Kill(true);
            await process.WaitForExitAsync();
            throw new TimeoutException($"PowerShell script '{scriptPath}' timed out.");
        }
        var result = new ProcessResult(process.ExitCode, stdout.ToString(), stderr.ToString());
        if (result.ExitCode != 0) { output.WriteLine(result.CombinedOutput); }
        return result;
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}

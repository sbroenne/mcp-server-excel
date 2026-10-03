using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PluginBootstrap")]
public sealed class PluginBootstrapTests(ITestOutputHelper output) : PluginTestBase(output)
{
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
    public void SourcePluginManifest_ConformsToAgentPluginsV1(string pluginName)
    {
        var pluginRoot = Path.Combine(RepoRoot, ".github", "plugins", pluginName);
        AssertAgentPluginManifest(pluginRoot, "0.0.0");
    }

    [Fact]
    public void ExcelMcpSource_UsesPortableNpxConfiguration()
    {
        AssertPortableMcpConfiguration(Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp"));
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
}

public abstract class PluginTestBase(ITestOutputHelper output)
{
    protected const string AgentPluginSchema = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json";
    protected const string AgentPluginMcpSchema = "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json";
    protected static readonly string RepoRoot = FindRepoRoot();
    protected static readonly string BuildPluginsScript = Path.Combine(RepoRoot, "scripts", "Build-Plugins.ps1");
    protected static readonly string SyncPublishedRepoScript = Path.Combine(RepoRoot, "scripts", "Sync-PublishedPluginRepo.ps1");

    [SupportedOSPlatform("windows")]
    protected static string[] SplitCommandLine(string commandLine)
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

    protected static void AssertAgentPluginManifest(string pluginRoot, string expectedVersion)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.Combine(pluginRoot, "plugin.json")));
        var root = document.RootElement;
        Assert.Equal(AgentPluginSchema, root.GetProperty("$schema").GetString());
        Assert.Equal(Path.GetFileName(pluginRoot), root.GetProperty("name").GetString());
        Assert.Equal(expectedVersion, root.GetProperty("version").GetString());
    }

    protected static void AssertPortableMcpConfiguration(string pluginRoot)
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

    protected static void AssertSkillDirectoryMatchesSource(
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

    protected static string CreateSandbox(string name)
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcp-{name}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(sandbox);
        return sandbox;
    }

    protected static void DeleteDirectoryIfExists(string path)
    {
        if (Directory.Exists(path)) { Directory.Delete(path, true); }
    }

    protected static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate repository root.");
    }

    protected async Task<ProcessResult> RunPowerShellFileAsync(
        string scriptPath,
        IReadOnlyList<string> arguments,
        Dictionary<string, string>? environmentVariables = null,
        int timeoutMs = 30000)
    {
        if (scriptPath == BuildPluginsScript && arguments.Contains("-Version", StringComparer.Ordinal)
            && !arguments.Contains("-SkillsDirectory", StringComparer.Ordinal))
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

    protected sealed record ProcessResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}

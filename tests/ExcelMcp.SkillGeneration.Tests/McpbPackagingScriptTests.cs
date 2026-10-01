using System.Diagnostics;
using System.IO.Compression;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

/// <summary>
/// Integration tests for MCPB packaging script behavior.
/// </summary>
[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
public sealed class McpbPackagingScriptTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string PackagingHelpers = Path.Combine(
        RepoRoot,
        "mcpb",
        "McpbPackaging.ps1");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Packaging")]
    public async Task AggregatePackaging_VsixTargetsRequireMatchingBundledRuntime(bool corruptArm64Payload)
    {
        var sandbox = CreateSandbox();
        try
        {
            var extension = Directory.CreateDirectory(Path.Combine(sandbox, "vscode-extension")).FullName;
            File.WriteAllText(Path.Combine(extension, "package.json"), """
                {"version":"1.2.3","extensionKind":["ui"],"os":["win32"],"scripts":{"vscode:prepublish":"npm run compile"}}
                """);
            File.WriteAllText(Path.Combine(sandbox, "CHANGELOG.md"), "Fixture changelog");
            var skills = Directory.CreateDirectory(Path.Combine(sandbox, "prepared", "excel-mcp")).FullName;
            File.WriteAllText(Path.Combine(skills, "VERSION"), "1.2.3");
            File.WriteAllText(Path.Combine(skills, "SKILL.md"), "Fixture skill");
            foreach (var (architecture, machine) in new[] { ("x64", (ushort)0x8664), ("arm64", (ushort)0xaa64) })
            {
                var runtime = Directory.CreateDirectory(Path.Combine(sandbox, "runtimes", architecture)).FullName;
                var header = new byte[70];
                header[0] = 0x4d;
                header[1] = 0x5a;
                BitConverter.GetBytes(64).CopyTo(header, 60);
                BitConverter.GetBytes(0x00004550).CopyTo(header, 64);
                BitConverter.GetBytes(machine).CopyTo(header, 68);
                File.WriteAllBytes(Path.Combine(runtime, "Sbroenne.ExcelMcp.McpServer.exe"), header);
            }
            var output = Directory.CreateDirectory(Path.Combine(sandbox, "artifacts", "packages")).FullName;
            var result = await RunPowerShellAsync($$"""
                $ErrorActionPreference = 'Stop'
                $root = '{{EscapePowerShellLiteral(sandbox)}}'
                $Version = '1.2.3'
                $SkillsDirectory = Join-Path $root 'prepared'
                $OutputDirectory = '{{EscapePowerShellLiteral(output)}}'
                $prepared = @{
                    Mcp = Join-Path $root 'runtimes\x64\Sbroenne.ExcelMcp.McpServer.exe'
                    'Mcp-arm64' = Join-Path $root 'runtimes\arm64\Sbroenne.ExcelMcp.McpServer.exe'
                }
                . '{{EscapePowerShellLiteral(Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1"))}}'
                $tokens = $null
                $errors = $null
                $ast = [Management.Automation.Language.Parser]::ParseFile(
                    '{{EscapePowerShellLiteral(Path.Combine(RepoRoot, "scripts", "Build-ReleasePackages.ps1"))}}',
                    [ref]$tokens, [ref]$errors)
                if ($errors.Count) { throw ($errors | Out-String) }
                foreach ($definition in $ast.FindAll({
                    param($node)
                    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                        $node.Name -in @('Invoke-PackageStep', 'Read-VsixEntry')
                }, $true)) {
                    . ([scriptblock]::Create($definition.Extent.Text))
                }
                function Publish-PackageRuntime { throw 'The prepared ARM64 runtime must be reused.' }
                function npm.cmd {
                    $global:LASTEXITCODE = 0
                    if ($args[0] -eq 'run' -and $args[1] -eq 'compile') {
                        New-Item -ItemType Directory -Path 'out' | Out-Null
                        Set-Content 'out\extension.js' 'fixture'
                        Set-Content 'out\prerequisites.js' 'fixture'
                    }
                    if ($args[0] -ne 'exec') { return }
                    $target = $args[[Array]::IndexOf($args, '--target') + 1]
                    $destination = $args[[Array]::IndexOf($args, '--out') + 1]
                    $archive = [IO.Compression.ZipFile]::Open($destination, [IO.Compression.ZipArchiveMode]::Create)
                    try {
                        foreach ($file in Get-ChildItem -File -Recurse) {
                            $relative = [IO.Path]::GetRelativePath((Get-Location).Path, $file.FullName).Replace('\', '/')
                            $source = $file.FullName
                            if (${{corruptArm64Payload.ToString().ToLowerInvariant()}} -and
                                $target -eq 'win32-arm64' -and $relative -eq 'bin/Sbroenne.ExcelMcp.McpServer.exe') {
                                $source = $prepared.Mcp
                            }
                            [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $source, "extension/$relative") | Out-Null
                        }
                        $writer = [IO.StreamWriter]::new($archive.CreateEntry('extension.vsixmanifest').Open())
                        try {
                            $writer.Write("<PackageManifest><Metadata><Identity TargetPlatform='$target'/></Metadata></PackageManifest>")
                        } finally { $writer.Dispose() }
                    } finally { $archive.Dispose() }
                }
                $extensionBlock = @($ast.FindAll({
                    param($node)
                    $node -is [Management.Automation.Language.IfStatementAst] -and
                        $node.Clauses[0].Item1.Extent.Text -eq '$Components -contains ''Extension'''
                }, $true))
                if ($extensionBlock.Count -ne 1) { throw 'Cannot locate the extension packaging phase.' }
                $body = $extensionBlock[0].Clauses[0].Item2.Extent.Text
                $extensionStage = $null
                try {
                    . ([scriptblock]::Create($body.Substring(1, $body.Length - 2)))
                } finally {
                    if ($extensionStage -and (Test-Path -LiteralPath $extensionStage)) {
                        Remove-Item -LiteralPath $extensionStage -Recurse -Force
                    }
                }
                """);
            if (corruptArm64Payload)
            {
                Assert.NotEqual(0, result.ExitCode);
                Assert.Contains("Runtime machine type 0x8664 does not match arm64", result.CombinedOutput, StringComparison.Ordinal);
            }
            else
            {
                Assert.True(result.ExitCode == 0, result.CombinedOutput);
                foreach (var (fileName, machine) in new[]
                {
                    ("excel-mcp-1.2.3.vsix", (ushort)0x8664),
                    ("excel-mcp-1.2.3-win32-arm64.vsix", (ushort)0xaa64)
                })
                {
                    using var archive = ZipFile.OpenRead(Path.Combine(output, fileName));
                    var entry = archive.GetEntry("extension/bin/Sbroenne.ExcelMcp.McpServer.exe");
                    Assert.NotNull(entry);
                    using var reader = new BinaryReader(entry.Open());
                    var header = reader.ReadBytes(70);
                    Assert.Equal(machine, BitConverter.ToUInt16(header, 68));
                }
                var debugRuntime = Path.Combine(output, "extension", "bin", "Sbroenne.ExcelMcp.McpServer.exe");
                var expectedDebugMachine = System.Runtime.InteropServices.RuntimeInformation.OSArchitecture ==
                    System.Runtime.InteropServices.Architecture.Arm64 ? (ushort)0xaa64 : (ushort)0x8664;
                Assert.Equal(expectedDebugMachine, BitConverter.ToUInt16(File.ReadAllBytes(debugRuntime), 68));
            }
            Assert.Empty(Directory.GetFiles(output, "*-server-inspection.exe"));
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Theory]
    [InlineData("wrong-version")]
    [InlineData("")]
    [Trait("Feature", "McpbPackaging")]
    public async Task AggregatePackaging_RejectsMismatchedPreparedSkillsBeforeCreatingOutput(string stamp)
    {
        var sandbox = CreateSandbox();
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
            foreach (var name in new[] { "Build-ReleasePackages.ps1", "Get-ValidationPlan.ps1", "PackageHelpers.ps1" })
            {
                File.Copy(Path.Combine(RepoRoot, "scripts", name), Path.Combine(scripts, name));
            }
            File.WriteAllText(Path.Combine(scripts, "Build-Plugins.ps1"), "throw 'Payload packaging was reached.'");
            var skills = Path.Combine(sandbox, "prepared");
            foreach (var name in new[] { "excel-cli", "excel-mcp" })
            {
                Directory.CreateDirectory(Path.Combine(skills, name));
                if (stamp.Length > 0) { File.WriteAllText(Path.Combine(skills, name, "VERSION"), stamp); }
            }
            var output = Path.Combine(sandbox, "artifacts", "packages");
            var result = await RunPowerShellAsync($$"""
                & '{{EscapePowerShellLiteral(Path.Combine(scripts, "Build-ReleasePackages.ps1"))}}' -Components Plugins -Version 1.2.3 -SkillsDirectory '{{EscapePowerShellLiteral(skills)}}' -OutputDirectory '{{EscapePowerShellLiteral(output)}}'
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("skill must match package version", result.Stderr, StringComparison.Ordinal);
            Assert.False(Directory.Exists(output));
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "McpbPackaging")]
    public async Task FailedBuild_PreservesExistingOutput()
    {
        var sandbox = CreateSandbox();
        try
        {
            var bundleRoot = Directory.CreateDirectory(Path.Combine(sandbox, "mcpb")).FullName;
            var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
            File.Copy(Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1"), Path.Combine(scripts, "PackageHelpers.ps1"));
            var builder = Path.Combine(bundleRoot, "Build-McpBundle.ps1");
            File.Copy(Path.Combine(RepoRoot, "mcpb", "Build-McpBundle.ps1"), builder);
            File.Copy(PackagingHelpers, Path.Combine(bundleRoot, "McpbPackaging.ps1"));
            var output = Path.Combine(bundleRoot, "artifacts");
            Directory.CreateDirectory(output);
            var previousPackage = Path.Combine(output, "excel-mcp-1.2.3.mcpb");
            var unrelatedFile = Path.Combine(output, "keep.txt");
            await File.WriteAllTextAsync(previousPackage, "previous-good-package");
            await File.WriteAllTextAsync(unrelatedFile, "unrelated");
            var result = await RunPowerShellAsync($$"""
                function dotnet { Write-Host 'publish-root-cause'; $global:LASTEXITCODE = 23 }
                & '{{EscapePowerShellLiteral(builder)}}' -Version '1.2.3'
                exit $LASTEXITCODE
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("publish-root-cause", result.Stdout + result.Stderr, StringComparison.Ordinal);
            Assert.True(File.Exists(previousPackage), "A failed build must preserve the previous package.");
            Assert.Equal("previous-good-package", await File.ReadAllTextAsync(previousPackage));
            Assert.Equal("unrelated", await File.ReadAllTextAsync(unrelatedFile));
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Packaging")]
    public async Task InstallPackageOutput_WhenInstallAndRestoreFail_PreservesBothErrorsAndRecoveryBackup()
    {
        var sandbox = CreateSandbox();
        try
        {
            var source = Directory.CreateDirectory(Path.Combine(sandbox, "source")).FullName;
            var destination = Directory.CreateDirectory(Path.Combine(sandbox, "destination")).FullName;
            await File.WriteAllTextAsync(Path.Combine(source, "payload.txt"), "replacement");
            await File.WriteAllTextAsync(Path.Combine(destination, "payload.txt"), "previous");

            var result = await RunPowerShellAsync($$"""
                $ErrorActionPreference = 'Stop'
                . '{{EscapePowerShellLiteral(Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1"))}}'
                $script:moveAttempt = 0
                function Move-Item {
                    param(
                        [Parameter(Mandatory)][string]$LiteralPath,
                        [Parameter(Mandatory)][string]$Destination
                    )
                    $script:moveAttempt++
                    if ($script:moveAttempt -eq 2) { throw 'installation root cause' }
                    if ($script:moveAttempt -eq 3) { throw 'restore root cause' }
                    Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination
                }
                function Remove-Item {
                    param(
                        [Parameter(Mandatory)][string]$LiteralPath,
                        [switch]$Recurse,
                        [switch]$Force
                    )
                    if ($LiteralPath.EndsWith('.tmp', [StringComparison]::OrdinalIgnoreCase)) {
                        throw 'temporary cleanup root cause'
                    }
                    Microsoft.PowerShell.Management\Remove-Item @PSBoundParameters
                }
                try {
                    Install-PackageOutput `
                        -Source '{{EscapePowerShellLiteral(source)}}' `
                        -Destination '{{EscapePowerShellLiteral(destination)}}'
                } catch {
                    [Console]::Error.WriteLine($_.Exception.Message)
                    exit 1
                }
                """);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("installation root cause", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("restore root cause", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("Recovery backup retained at", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("temporary cleanup root cause", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("Temporary output retained at", result.CombinedOutput, StringComparison.Ordinal);
            Assert.False(Directory.Exists(destination));

            var backup = Assert.Single(Directory.GetDirectories(sandbox, "*.bak"));
            Assert.Equal("previous", await File.ReadAllTextAsync(Path.Combine(backup, "payload.txt")));
            Assert.Contains(backup, result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
            var temporary = Assert.Single(Directory.GetDirectories(sandbox, "*.tmp"));
            Assert.Equal("replacement", await File.ReadAllTextAsync(Path.Combine(temporary, "payload.txt")));
            Assert.Contains(temporary, result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Packaging")]
    public async Task InstallPackageOutput_WhenBackupCleanupFails_ReportsSuccessAndRetainedBackup()
    {
        var sandbox = CreateSandbox();
        try
        {
            var source = Directory.CreateDirectory(Path.Combine(sandbox, "source")).FullName;
            var destination = Directory.CreateDirectory(Path.Combine(sandbox, "destination")).FullName;
            await File.WriteAllTextAsync(Path.Combine(source, "payload.txt"), "replacement");
            await File.WriteAllTextAsync(Path.Combine(destination, "payload.txt"), "previous");

            var result = await RunPowerShellAsync($$"""
                $ErrorActionPreference = 'Stop'
                . '{{EscapePowerShellLiteral(Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1"))}}'
                function Remove-Item {
                    param(
                        [Parameter(Mandatory)][string]$LiteralPath,
                        [switch]$Recurse,
                        [switch]$Force
                    )
                    if ($LiteralPath.EndsWith('.bak', [StringComparison]::OrdinalIgnoreCase)) {
                        throw 'backup cleanup root cause'
                    }
                    Microsoft.PowerShell.Management\Remove-Item @PSBoundParameters
                }
                Install-PackageOutput `
                    -Source '{{EscapePowerShellLiteral(source)}}' `
                    -Destination '{{EscapePowerShellLiteral(destination)}}'
                Write-Output 'install returned successfully'
                """);

            Assert.Equal(0, result.ExitCode);
            Assert.Contains("install returned successfully", result.Stdout, StringComparison.Ordinal);
            Assert.Contains("installed successfully", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("backup cleanup root cause", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("Backup retained at", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Equal("replacement", await File.ReadAllTextAsync(Path.Combine(destination, "payload.txt")));

            var backup = Assert.Single(Directory.GetDirectories(sandbox, "*.bak"));
            Assert.Equal("previous", await File.ReadAllTextAsync(Path.Combine(backup, "payload.txt")));
            Assert.Contains(backup, result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Theory]
    [InlineData("src/output")]
    [InlineData("skills/generated")]
    [InlineData(".github/plugins/output")]
    [Trait("Feature", "McpbPackaging")]
    public async Task PackageOutputs_RejectSourceDestinations(string relativePath)
    {
        var sandbox = CreateSandbox();
        try
        {
            var output = Path.Combine(sandbox, relativePath);
            var result = await RunPowerShellAsync($$"""
                . '{{EscapePowerShellLiteral(Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1"))}}'
                Assert-PackageOutputPath -Path '{{EscapePowerShellLiteral(output)}}' -RepoRoot '{{EscapePowerShellLiteral(sandbox)}}'
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Unsafe package output", result.Stderr, StringComparison.Ordinal);
            Assert.False(Directory.Exists(output));
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "McpbPackaging")]
    public async Task PublicationSync_RejectsItsOwnSourceTreeBeforeInspectingPayload()
    {
        var sandbox = CreateSandbox();
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
            var built = Directory.CreateDirectory(Path.Combine(sandbox, "built")).FullName;
            foreach (var name in new[] { "Sync-PublishedPluginRepo.ps1", "PackageHelpers.ps1" })
            {
                File.Copy(Path.Combine(RepoRoot, "scripts", name), Path.Combine(scripts, name));
            }
            var result = await RunPowerShellAsync($$"""
                & '{{EscapePowerShellLiteral(Path.Combine(scripts, "Sync-PublishedPluginRepo.ps1"))}}' -PublishedRepoDir '{{EscapePowerShellLiteral(sandbox)}}' -BuiltPluginsDir '{{EscapePowerShellLiteral(built)}}' -Version 1.2.3
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Unsafe package output", result.Stderr, StringComparison.Ordinal);
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Theory]
    [InlineData("version.txt")]
    [InlineData("skills/excel-cli/SKILL.md")]
    [Trait("Feature", "McpbPackaging")]
    public async Task PublicationSync_IncompletePayloadCannotReplaceExistingOutput(string missingFile)
    {
        var sandbox = CreateSandbox();
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
            foreach (var name in new[] { "Sync-PublishedPluginRepo.ps1", "PackageHelpers.ps1" })
            {
                File.Copy(Path.Combine(RepoRoot, "scripts", name), Path.Combine(scripts, name));
            }
            var overlay = Directory.CreateDirectory(Path.Combine(sandbox, ".github", "plugins", "marketplace-repo")).FullName;
            File.WriteAllText(Path.Combine(overlay, "README.md"), "new overlay");
            var built = Path.Combine(sandbox, "artifacts", "built");
            foreach (var name in new[] { "excel-cli", "excel-mcp" })
            {
                var plugin = Directory.CreateDirectory(Path.Combine(built, name)).FullName;
                var manifest = File.ReadAllText(Path.Combine(RepoRoot, ".github", "plugins", name, "plugin.json"));
                File.WriteAllText(Path.Combine(plugin, "plugin.json"), manifest.Replace("0.0.0", "1.2.3", StringComparison.Ordinal));
                foreach (var file in new[]
                {
                    "README.md", "version.txt", $"skills/{name}/SKILL.md", $"skills/{name}/VERSION",
                    $"skills/{name}/references/range.md",
                })
                {
                    var destination = Path.Combine(plugin, file);
                    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                    File.WriteAllText(destination, "1.2.3");
                }
                if (name == "excel-cli")
                {
                    var wrapper = Path.Combine(plugin, "bin", "start-cli.ps1");
                    Directory.CreateDirectory(Path.GetDirectoryName(wrapper)!);
                    File.WriteAllText(wrapper, "1.2.3");
                }
                else
                {
                    File.Copy(
                        Path.Combine(RepoRoot, ".github", "plugins", "excel-mcp", "mcp.json"),
                        Path.Combine(plugin, "mcp.json"));
                }
            }
            File.Delete(Path.Combine(built, "excel-cli", missingFile));
            var output = Directory.CreateDirectory(Path.Combine(sandbox, "artifacts", "published")).FullName;
            File.WriteAllText(Path.Combine(output, "README.md"), "previous output");
            var result = await RunPowerShellAsync($$"""
                & '{{EscapePowerShellLiteral(Path.Combine(scripts, "Sync-PublishedPluginRepo.ps1"))}}' -PublishedRepoDir '{{EscapePowerShellLiteral(output)}}' -BuiltPluginsDir '{{EscapePowerShellLiteral(built)}}' -Version 1.2.3
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Incomplete plugin payload", result.Stderr, StringComparison.Ordinal);
            Assert.Equal("previous output", File.ReadAllText(Path.Combine(output, "README.md")));
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "McpbPackaging")]
    public async Task RemoveStagingDirectory_RetriesTransientFileLockUntilDirectoryIsGone()
    {
        var sandbox = CreateSandbox();
        try
        {
            var script = $$"""
                $ErrorActionPreference = 'Stop'
                . '{{EscapePowerShellLiteral(PackagingHelpers)}}'
                $script:attempts = 0
                $removeDirectory = {
                    param([string]$Path)
                    $script:attempts++
                    if ($script:attempts -lt 3) {
                        throw [System.IO.IOException]::new('locked by scanner')
                    }

                    [System.IO.Directory]::Delete($Path, $true)
                }

                Remove-McpbStagingDirectory `
                    -Path '{{EscapePowerShellLiteral(sandbox)}}' `
                    -Timeout ([TimeSpan]::FromSeconds(1)) `
                    -RetryInterval ([TimeSpan]::Zero) `
                    -RemoveDirectory $removeDirectory
                Write-Output "attempts=$script:attempts"
                """;

            var result = await RunPowerShellAsync(script);

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("attempts=3", result.Stdout, StringComparison.Ordinal);
            Assert.False(Directory.Exists(sandbox));
        }
        finally
        {
            if (Directory.Exists(sandbox))
            {
                Directory.Delete(sandbox, recursive: true);
            }
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "McpbPackaging")]
    public async Task RemoveStagingDirectory_WhenTimeoutExpires_ReportsTerminalState()
    {
        var sandbox = CreateSandbox();
        try
        {
            var script = $$"""
                $ErrorActionPreference = 'Stop'
                . '{{EscapePowerShellLiteral(PackagingHelpers)}}'
                $removeDirectory = {
                    param([string]$Path)
                    throw [System.UnauthorizedAccessException]::new('locked by scanner')
                }

                Remove-McpbStagingDirectory `
                    -Path '{{EscapePowerShellLiteral(sandbox)}}' `
                    -Timeout ([TimeSpan]::Zero) `
                    -RetryInterval ([TimeSpan]::Zero) `
                    -RemoveDirectory $removeDirectory
                """;

            var result = await RunPowerShellAsync(script);

            Assert.NotEqual(0, result.ExitCode);
            Assert.True(Directory.Exists(sandbox));
            Assert.Contains(
                "Failed to remove MCPB staging directory",
                result.Stderr,
                StringComparison.Ordinal);
            Assert.Contains(sandbox, result.Stderr, StringComparison.Ordinal);
            Assert.Contains("after 1 attempt", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("within 0 ms", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("stale staging remains", result.Stderr, StringComparison.Ordinal);
            Assert.Contains("locked by scanner", result.Stderr, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
        }
    }

    private static string CreateSandbox()
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcpMcpbPackaging-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path.Combine(sandbox, "server"));
        File.WriteAllText(Path.Combine(sandbox, "server", "excel-mcp-server.exe"), "test");
        return sandbox;
    }

    private static string EscapePowerShellLiteral(string value) =>
        value.Replace("'", "''", StringComparison.Ordinal);

    private static async Task<ScriptResult> RunPowerShellAsync(string script)
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
        startInfo.ArgumentList.Add("-Command");
        startInfo.ArgumentList.Add(script);

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);

        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw;
        }

        return new ScriptResult(process.ExitCode, await stdout, await stderr);
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

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}

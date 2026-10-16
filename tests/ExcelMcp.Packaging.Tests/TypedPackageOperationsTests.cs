using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "Packaging")]
public sealed class TypedPackageOperationsTests
{
    private static string Helpers => Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1");
    private static string ExtensionHelpers => Path.Combine(RepoRoot, "scripts", "ExtensionPackaging.ps1");

    [Theory]
    [InlineData("Cli", "x64", "ExcelMcp-CLI-1.2.3-windows.zip", "excelcli.exe", 0x8664)]
    [InlineData("Cli", "arm64", "ExcelMcp-CLI-1.2.3-windows-arm64.zip", "excelcli.exe", 0xaa64)]
    [InlineData("Mcp", "x64", "ExcelMcp-MCP-Server-1.2.3-windows.zip", "mcp-excel.exe", 0x8664)]
    [InlineData("Mcp", "arm64", "ExcelMcp-MCP-Server-1.2.3-windows-arm64.zip", "mcp-excel.exe", 0xaa64)]
    public async Task StandaloneArchive_ContainsMatchingArchitectureRuntimeAndSupportFiles(
        string component, string architecture, string expectedName, string executableName, int expectedMachine)
    {
        var sandbox = CreateSandbox("standalone");
        try
        {
            var output = Directory.CreateDirectory(Path.Combine(sandbox, "artifacts")).FullName;
            var payload = File.ReadAllBytes(typeof(TypedPackageOperationsTests).Assembly.Location);
            var offset = BitConverter.ToInt32(payload, 0x3c);
            BitConverter.GetBytes((ushort)expectedMachine).CopyTo(payload, offset + 4);
            var executable = Path.Combine(sandbox, executableName);
            File.WriteAllBytes(executable, payload);
            foreach (var file in new[] { "README.md", "LICENSE", "CHANGELOG.md" }) { File.WriteAllText(Path.Combine(sandbox, file), file); }
            var result = await RunPowerShellCommandAsync($"""
                . '{Helpers}'
                New-StandalonePackage '{component}' '1.2.3' '{executable}' '{architecture}' '{output}' '{sandbox}'
                """);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            using var zip = ZipFile.OpenRead(Path.Combine(output, expectedName));
            foreach (var file in new[] { executableName, "README.md", "LICENSE", "CHANGELOG.md" }) { Assert.NotNull(zip.GetEntry(file)); }
            using var stream = zip.GetEntry(executableName)!.Open();
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            Assert.Equal((ushort)expectedMachine, BitConverter.ToUInt16(memory.ToArray(), offset + 4));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Theory]
    [InlineData(false, null, false)]
    [InlineData(true, null, false)]
    [InlineData(false, "AGENTS.md", false)]
    [InlineData(false, "CLAUDE.md", false)]
    [InlineData(false, null, true)]
    [InlineData(true, null, true)]
    public async Task ExtensionPackages_CompileOnceAndRequireMatchingBundledRuntime(
        bool corruptArm64Payload, string? developerFile, bool skipExtensionTests)
    {
        var sandbox = CreateSandbox("vsix");
        try
        {
            var version = FileVersionInfo.GetVersionInfo(typeof(TypedPackageOperationsTests).Assembly.Location).ProductVersion!.Split('+')[0];
            var payload = File.ReadAllBytes(typeof(TypedPackageOperationsTests).Assembly.Location);
            var offset = BitConverter.ToInt32(payload, 0x3c);
            var extension = Directory.CreateDirectory(Path.Combine(sandbox, "vscode-extension")).FullName;
            if (developerFile is not null) { File.WriteAllText(Path.Combine(extension, developerFile), "Developer-only instructions"); }
            File.WriteAllText(Path.Combine(extension, "package.json"), $$$"""
                {"version":"{{{version}}}","extensionKind":["ui"],"os":["win32"],"scripts":{"vscode:prepublish":"npm run compile"}}
                """);
            File.WriteAllText(Path.Combine(sandbox, "CHANGELOG.md"), "Fixture changelog");
            var skills = Path.Combine(sandbox, "prepared");
            var skill = Directory.CreateDirectory(Path.Combine(skills, "excel-mcp-report-formatting")).FullName;
            File.WriteAllText(Path.Combine(skill, "VERSION"), version);
            File.WriteAllText(Path.Combine(skill, "SKILL.md"), "Fixture skill");
            foreach (var (architecture, machine) in new[] { ("x64", (ushort)0x8664), ("arm64", (ushort)0xaa64) })
            {
                var directory = Directory.CreateDirectory(Path.Combine(sandbox, "runtimes", architecture)).FullName;
                BitConverter.GetBytes(machine).CopyTo(payload, offset + 4);
                File.WriteAllBytes(Path.Combine(directory, "Sbroenne.ExcelMcp.McpServer.exe"), payload);
            }
            var output = Directory.CreateDirectory(Path.Combine(sandbox, "artifacts", "packages")).FullName;
            var result = await RunPowerShellCommandAsync($$$"""
                . '{{{Helpers}}}'
                . '{{{ExtensionHelpers}}}'
                $root = '{{{sandbox}}}'
                $calls = [Collections.Generic.List[string]]::new()
                function Invoke-PackageCommand {
                    param($Directory, $Executable, $Arguments)
                    if ($Executable -ne 'npm') { throw 'Prepared runtimes must not be rebuilt.' }
                    $calls.Add($Arguments -join ' ')
                    if ($Arguments -join ' ' -eq 'run compile') {
                        New-Item "$Directory\out" -ItemType Directory -Force | Out-Null
                        'fixture' | Set-Content "$Directory\out\extension.js"
                        'fixture' | Set-Content "$Directory\out\prerequisites.js"
                    }
                    if ($Arguments[0] -eq 'exec') {
                        $manifest = Read-PackageJson "$Directory\package.json"
                        if ($manifest.scripts.Contains('vscode:prepublish')) { throw 'VSIX must reuse compiled source.' }
                        $target = $Arguments[[Array]::IndexOf($Arguments, '--target') + 1]
                        $destination = $Arguments[[Array]::IndexOf($Arguments, '--out') + 1]
                        $zip = [IO.Compression.ZipFile]::Open($destination, [IO.Compression.ZipArchiveMode]::Create)
                        try {
                            foreach ($file in [IO.Directory]::EnumerateFiles($Directory, '*', [IO.SearchOption]::AllDirectories)) {
                                $relative = [IO.Path]::GetRelativePath($Directory, $file).Replace('\', '/')
                                $source = $file
                                if (${{{corruptArm64Payload}}} -and $target -eq 'win32-arm64' -and $relative -eq 'bin/Sbroenne.ExcelMcp.McpServer.exe') {
                                    $source = "$root\runtimes\x64\Sbroenne.ExcelMcp.McpServer.exe"
                                }
                                [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $source, "extension/$relative") | Out-Null
                            }
                            $writer = [IO.StreamWriter]::new($zip.CreateEntry('extension.vsixmanifest').Open())
                            try { $writer.Write("<PackageManifest><Metadata><Identity TargetPlatform='$target'/></Metadata></PackageManifest>") }
                            finally { $writer.Dispose() }
                        } finally { $zip.Dispose() }
                    }
                }
                $stage = Initialize-ExtensionPackage $root '{{{version}}}' '{{{skills}}}' -SkipExtensionTests:${{{skipExtensionTests}}}
                try {
                    Complete-ExtensionPackages $stage '{{{version}}}' '{{{output}}}' @{
                        Mcp = "$root\runtimes\x64\Sbroenne.ExcelMcp.McpServer.exe"
                        'Mcp-arm64' = "$root\runtimes\arm64\Sbroenne.ExcelMcp.McpServer.exe"
                    }
                } finally {
                    $calls | ConvertTo-Json | Set-Content "$root\calls.json"
                    Remove-PackageStaging $stage
                }
                """);
            var calls = System.Text.Json.JsonSerializer.Deserialize<string[]>(File.ReadAllText(Path.Combine(sandbox, "calls.json")))!;
            Assert.Contains("ci --ignore-scripts", calls);
            Assert.Single(calls, call => call == "run compile");
            Assert.Contains("run lint", calls);
            Assert.Equal(!skipExtensionTests, calls.Contains("run typecheck:tests"));
            Assert.Equal(!skipExtensionTests, calls.Contains("test"));
            if (developerFile is not null)
            {
                Assert.NotEqual(0, result.ExitCode);
                Assert.Contains($"VSIX contains development files or the CLI: extension/{developerFile}", result.CombinedOutput, StringComparison.Ordinal);
            }
            else if (corruptArm64Payload)
            {
                Assert.NotEqual(0, result.ExitCode);
                Assert.Contains("Runtime machine type 0x8664 does not match arm64", result.CombinedOutput, StringComparison.Ordinal);
            }
            else
            {
                Assert.True(result.ExitCode == 0, result.CombinedOutput);
                foreach (var (architecture, filename, machine) in new[] {
                    ("x64", $"excel-mcp-{version}.vsix", (ushort)0x8664),
                    ("arm64", $"excel-mcp-{version}-win32-arm64.vsix", (ushort)0xaa64)
                })
                {
                    using var archive = ZipFile.OpenRead(Path.Combine(output, filename));
                    var entry = archive.GetEntry("extension/bin/Sbroenne.ExcelMcp.McpServer.exe");
                    Assert.NotNull(entry);
                    using var reader = new BinaryReader(entry.Open());
                    var expected = File.ReadAllBytes(Path.Combine(sandbox, "runtimes", architecture, "Sbroenne.ExcelMcp.McpServer.exe"));
                    var actual = reader.ReadBytes(expected.Length);
                    Assert.Equal(machine, BitConverter.ToUInt16(actual, offset + 4));
                    Assert.Equal(expected, actual);
                    using var skillReader = new StreamReader(archive.GetEntry("extension/skills/excel-mcp-report-formatting/SKILL.md")!.Open());
                    Assert.Equal("Fixture skill", skillReader.ReadToEnd());
                }
                var debug = File.ReadAllBytes(Path.Combine(output, "extension", "bin", "Sbroenne.ExcelMcp.McpServer.exe"));
                Assert.Equal(RuntimeInformation.OSArchitecture == Architecture.Arm64 ? (ushort)0xaa64 : (ushort)0x8664, BitConverter.ToUInt16(debug, offset + 4));
            }
            Assert.Empty(Directory.GetFiles(output, "*-server-inspection.exe"));
            Assert.Contains("vscode:prepublish", File.ReadAllText(Path.Combine(extension, "package.json")), StringComparison.Ordinal);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task InstallPackageOutput_WhenInstallAndRestoreFail_PreservesBothErrorsAndRecoveryBackup()
    {
        var sandbox = CreateSandbox("install-recovery");
        try
        {
            var source = Directory.CreateDirectory(Path.Combine(sandbox, "source")).FullName;
            var destination = Directory.CreateDirectory(Path.Combine(sandbox, "destination")).FullName;
            File.WriteAllText(Path.Combine(source, "payload.txt"), "replacement");
            File.WriteAllText(Path.Combine(destination, "payload.txt"), "previous");
            var result = await RunPowerShellCommandAsync($$"""
                . '{{Helpers}}'
                $script:attempts = 0
                Install-PackageOutput '{{source}}' '{{destination}}' -Move {
                    param($from, $to)
                    $script:attempts++
                    if ($script:attempts -eq 2) { throw [IO.IOException]::new('installation root cause') }
                    if ($script:attempts -eq 3) { throw [IO.IOException]::new('restore root cause') }
                    Move-PackageItem $from $to
                } -Delete {
                    param($path)
                    if ($path.EndsWith('.tmp')) { throw [IO.IOException]::new('temporary cleanup root cause') }
                    Remove-PackageItem $path
                }
                """);
            Assert.NotEqual(0, result.ExitCode);
            foreach (var text in new[] { "installation root cause", "restore root cause", "Recovery backup retained at", "temporary cleanup root cause", "Temporary output retained at" })
            {
                Assert.Contains(text, result.CombinedOutput, StringComparison.Ordinal);
            }
            Assert.False(Directory.Exists(destination));
            var backup = Assert.Single(Directory.GetDirectories(sandbox, "*.bak"));
            Assert.Equal("previous", File.ReadAllText(Path.Combine(backup, "payload.txt")));
            var temporary = Assert.Single(Directory.GetDirectories(sandbox, "*.tmp"));
            Assert.Equal("replacement", File.ReadAllText(Path.Combine(temporary, "payload.txt")));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    public async Task InstallPackageOutput_WhenBackupCleanupFails_ReportsSuccessAndRetainedBackup()
    {
        var sandbox = CreateSandbox("install-cleanup");
        try
        {
            var source = Directory.CreateDirectory(Path.Combine(sandbox, "source")).FullName;
            var destination = Directory.CreateDirectory(Path.Combine(sandbox, "destination")).FullName;
            File.WriteAllText(Path.Combine(source, "payload.txt"), "replacement");
            File.WriteAllText(Path.Combine(destination, "payload.txt"), "previous");
            var result = await RunPowerShellCommandAsync($$"""
                . '{{Helpers}}'
                Install-PackageOutput '{{source}}' '{{destination}}' -Delete {
                    param($path)
                    if ($path.EndsWith('.bak')) { throw [IO.IOException]::new('backup cleanup root cause') }
                    Remove-PackageItem $path
                }
                """);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("installed successfully", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("backup cleanup root cause", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Contains("Backup retained at", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Equal("replacement", File.ReadAllText(Path.Combine(destination, "payload.txt")));
            var backup = Assert.Single(Directory.GetDirectories(sandbox, "*.bak"));
            Assert.Equal("previous", File.ReadAllText(Path.Combine(backup, "payload.txt")));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    [Trait("Feature", "McpbPackaging")]
    public async Task RemoveStagingDirectory_RetriesTransientFileLockUntilDirectoryIsGone()
    {
        var sandbox = CreateSandbox("staging-retry");
        try
        {
            var result = await RunPowerShellCommandAsync($$"""
                . '{{Helpers}}'
                $script:attempts = 0
                Remove-PackageStaging '{{sandbox}}' -TimeoutMilliseconds 1000 -RetryMilliseconds 0 -Remove {
                    param($path)
                    $script:attempts++
                    if ($script:attempts -lt 3) { throw [IO.IOException]::new('locked by scanner') }
                    Remove-PackageItem $path
                }
                if ($script:attempts -ne 3) { throw 'Unexpected retry count.' }
                """);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.False(Directory.Exists(sandbox));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact]
    [Trait("Feature", "McpbPackaging")]
    public async Task RemoveStagingDirectory_WhenTimeoutExpires_ReportsTerminalState()
    {
        var sandbox = CreateSandbox("staging-timeout");
        try
        {
            var result = await RunPowerShellCommandAsync($$"""
                . '{{Helpers}}'
                try {
                    Remove-PackageStaging '{{sandbox}}' -TimeoutMilliseconds 0 -RetryMilliseconds 0 -Remove {
                        throw [UnauthorizedAccessException]::new('locked by scanner')
                    }
                } catch [IO.IOException] {
                    [Console]::Error.WriteLine($_.Exception.Message)
                    throw
                }
                """);
            Assert.NotEqual(0, result.ExitCode);
            Assert.True(Directory.Exists(sandbox));
            foreach (var text in new[] { "Failed to remove package staging directory", sandbox, "after 1 attempt", "within 0 ms", "Stale staging remains", "locked by scanner" })
            {
                Assert.Contains(text, result.CombinedOutput, StringComparison.Ordinal);
            }
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }
}

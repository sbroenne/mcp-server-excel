using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "Packaging")]
public sealed class TypedPackageOperationsTests
{
    [Theory]
    [InlineData(false, true, null, false)]
    [InlineData(true, true, null, false)]
    [InlineData(false, false, null, false)]
    [InlineData(true, false, null, false)]
    [InlineData(false, true, "AGENTS.md", false)]
    [InlineData(false, true, "CLAUDE.md", false)]
    [InlineData(false, true, null, true)]
    [InlineData(true, true, null, true)]
    public async Task AggregatePackaging_VsixTargetsRequireMatchingBundledRuntime(
        bool corruptArm64Payload, bool reuseArm64Runtime, string? developerFile, bool skipExtensionTests)
    {
        var sandbox = NewSandbox();
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
            var prepared = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var (architecture, machine) in new[] { ("x64", (ushort)0x8664), ("arm64", (ushort)0xaa64) })
            {
                var directory = Directory.CreateDirectory(Path.Combine(sandbox, "runtimes", architecture)).FullName;
                BitConverter.GetBytes(machine).CopyTo(payload, offset + 4);
                File.WriteAllBytes(Path.Combine(directory, "Sbroenne.ExcelMcp.McpServer.exe"), payload);
            }
            prepared["Mcp"] = Path.Combine(sandbox, "runtimes", "x64", "Sbroenne.ExcelMcp.McpServer.exe");
            if (reuseArm64Runtime) { prepared["Mcp-arm64"] = Path.Combine(sandbox, "runtimes", "arm64", "Sbroenne.ExcelMcp.McpServer.exe"); }
            var output = Directory.CreateDirectory(Path.Combine(sandbox, "artifacts", "packages")).FullName;
            var calls = new List<PackageCommand>();
            var publications = 0;
            var commands = new PackageCommands(call =>
            {
                calls.Add(call);
                if (call.Executable == "dotnet")
                {
                    Assert.False(reuseArm64Runtime, "The prepared ARM64 runtime must be reused.");
                    publications++;
                    Assert.Equal(sandbox, call.WorkingDirectory);
                    Assert.Contains(Path.Combine(sandbox, "src", "ExcelMcp.McpServer", "ExcelMcp.McpServer.csproj"), call.Arguments);
                    Assert.Contains("win-arm64", call.Arguments);
                    Assert.Contains($"-p:Version={version}", call.Arguments);
                    var destination = call.Arguments[Array.IndexOf(call.Arguments, "-o") + 1];
                    Assert.Equal(Path.Combine(output, "runtimes", "Mcp-arm64"), destination);
                    Directory.CreateDirectory(destination);
                    File.Copy(Path.Combine(sandbox, "runtimes", "arm64", "Sbroenne.ExcelMcp.McpServer.exe"), Path.Combine(destination, "Sbroenne.ExcelMcp.McpServer.exe"));
                }
                else
                {
                    Assert.Equal("npm", call.Executable);
                    if (call.Arguments is ["run", "compile"])
                    {
                        var compiled = Directory.CreateDirectory(Path.Combine(call.WorkingDirectory, "out")).FullName;
                        File.WriteAllText(Path.Combine(compiled, "extension.js"), "fixture");
                        File.WriteAllText(Path.Combine(compiled, "prerequisites.js"), "fixture");
                    }
                    if (call.Arguments[0] == "exec")
                    {
                        var target = call.Arguments[Array.IndexOf(call.Arguments, "--target") + 1];
                        var destination = call.Arguments[Array.IndexOf(call.Arguments, "--out") + 1];
                        using var archive = ZipFile.Open(destination, ZipArchiveMode.Create);
                        foreach (var file in Directory.EnumerateFiles(call.WorkingDirectory, "*", SearchOption.AllDirectories))
                        {
                            var relative = Path.GetRelativePath(call.WorkingDirectory, file).Replace('\\', '/');
                            var source = corruptArm64Payload && target == "win32-arm64" && relative == "bin/Sbroenne.ExcelMcp.McpServer.exe" ? prepared["Mcp"] : file;
                            archive.CreateEntryFromFile(source, $"extension/{relative}");
                        }
                        using var writer = new StreamWriter(archive.CreateEntry("extension.vsixmanifest").Open());
                        writer.Write($"<PackageManifest><Metadata><Identity TargetPlatform='{target}'/></Metadata></PackageManifest>");
                    }
                }
                return Task.FromResult(new ProcessResult(0, "", ""));
            });
            var failure = await Record.ExceptionAsync(() => new PackageExecution(sandbox, commands).ExtensionAsync(version, skills, output, prepared, skipExtensionTests));
            Assert.Contains(calls, call => call.Arguments.SequenceEqual(["ci", "--ignore-scripts"]));
            Assert.Contains(calls, call => call.Arguments.SequenceEqual(["run", "compile"]));
            Assert.Contains(calls, call => call.Arguments.SequenceEqual(["run", "lint"]));
            Assert.Equal(!skipExtensionTests, calls.Any(call => call.Arguments.SequenceEqual(["run", "typecheck:tests"])));
            Assert.Equal(!skipExtensionTests, calls.Any(call => call.Arguments.SequenceEqual(["test"])));
            if (developerFile is not null)
            {
                Assert.NotNull(failure);
                Assert.Contains($"VSIX contains development files or the CLI: extension/{developerFile}", failure.Message, StringComparison.Ordinal);
            }
            else if (corruptArm64Payload)
            {
                Assert.NotNull(failure);
                Assert.Contains("Runtime machine type 0x8664 does not match arm64", failure.Message, StringComparison.Ordinal);
            }
            else
            {
                Assert.Null(failure);
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
                }
                var debug = File.ReadAllBytes(Path.Combine(output, "extension", "bin", "Sbroenne.ExcelMcp.McpServer.exe"));
                Assert.Equal(RuntimeInformation.OSArchitecture == Architecture.Arm64 ? (ushort)0xaa64 : (ushort)0x8664, BitConverter.ToUInt16(debug, offset + 4));
            }
            Assert.Equal(reuseArm64Runtime ? 0 : 1, publications);
            if (!reuseArm64Runtime)
            {
                Assert.Equal(payload, File.ReadAllBytes(Path.Combine(output, "runtimes", "Mcp-arm64", "Sbroenne.ExcelMcp.McpServer.exe")));
            }
            Assert.Empty(Directory.GetFiles(output, "*-server-inspection.exe"));
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    public void InstallPackageOutput_WhenInstallAndRestoreFail_PreservesBothErrorsAndRecoveryBackup()
    {
        var sandbox = NewSandbox();
        try
        {
            var source = Directory.CreateDirectory(Path.Combine(sandbox, "source")).FullName;
            var destination = Directory.CreateDirectory(Path.Combine(sandbox, "destination")).FullName;
            File.WriteAllText(Path.Combine(source, "payload.txt"), "replacement");
            File.WriteAllText(Path.Combine(destination, "payload.txt"), "previous");
            var attempts = 0;
            var warnings = new List<string>();
            var error = Assert.Throws<AggregateException>(() => PackageFiles.Install(source, destination,
                move: (from, to) =>
                {
                    attempts++;
                    if (attempts == 2) { throw new IOException("installation root cause"); }
                    if (attempts == 3) { throw new IOException("restore root cause"); }
                    Directory.Move(from, to);
                },
                delete: path =>
                {
                    if (path.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase)) { throw new IOException("temporary cleanup root cause"); }
                    Directory.Delete(path, recursive: true);
                },
                warning: warnings.Add));
            var messages = error.Message + "\n" + string.Join('\n', warnings);
            Assert.Contains("installation root cause", messages, StringComparison.Ordinal);
            Assert.Contains("restore root cause", messages, StringComparison.Ordinal);
            Assert.Contains("Recovery backup retained at", messages, StringComparison.Ordinal);
            Assert.Contains("temporary cleanup root cause", messages, StringComparison.Ordinal);
            Assert.Contains("Temporary output retained at", messages, StringComparison.Ordinal);
            Assert.False(Directory.Exists(destination));
            var backup = Assert.Single(Directory.GetDirectories(sandbox, "*.bak"));
            Assert.Equal("previous", File.ReadAllText(Path.Combine(backup, "payload.txt")));
            Assert.Contains(backup, messages, StringComparison.OrdinalIgnoreCase);
            var temporary = Assert.Single(Directory.GetDirectories(sandbox, "*.tmp"));
            Assert.Equal("replacement", File.ReadAllText(Path.Combine(temporary, "payload.txt")));
            Assert.Contains(temporary, messages, StringComparison.OrdinalIgnoreCase);
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    public void InstallPackageOutput_WhenBackupCleanupFails_ReportsSuccessAndRetainedBackup()
    {
        var sandbox = NewSandbox();
        try
        {
            var source = Directory.CreateDirectory(Path.Combine(sandbox, "source")).FullName;
            var destination = Directory.CreateDirectory(Path.Combine(sandbox, "destination")).FullName;
            File.WriteAllText(Path.Combine(source, "payload.txt"), "replacement");
            File.WriteAllText(Path.Combine(destination, "payload.txt"), "previous");
            var warnings = new List<string>();
            PackageFiles.Install(source, destination, delete: path =>
            {
                if (path.EndsWith(".bak", StringComparison.OrdinalIgnoreCase)) { throw new IOException("backup cleanup root cause"); }
                Directory.Delete(path, recursive: true);
            }, warning: warnings.Add);
            var messages = string.Join('\n', warnings);
            Assert.Contains("installed successfully", messages, StringComparison.Ordinal);
            Assert.Contains("backup cleanup root cause", messages, StringComparison.Ordinal);
            Assert.Contains("Backup retained at", messages, StringComparison.Ordinal);
            Assert.Equal("replacement", File.ReadAllText(Path.Combine(destination, "payload.txt")));
            var backup = Assert.Single(Directory.GetDirectories(sandbox, "*.bak"));
            Assert.Equal("previous", File.ReadAllText(Path.Combine(backup, "payload.txt")));
            Assert.Contains(backup, messages, StringComparison.OrdinalIgnoreCase);
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Fact]
    [Trait("Feature", "McpbPackaging")]
    public void RemoveStagingDirectory_RetriesTransientFileLockUntilDirectoryIsGone()
    {
        var sandbox = NewSandbox();
        try
        {
            var attempts = 0;
            PackageFiles.RemoveStaging(sandbox, TimeSpan.FromSeconds(1), TimeSpan.Zero, path =>
            {
                attempts++;
                if (attempts < 3) { throw new IOException("locked by scanner"); }
                Directory.Delete(path, recursive: true);
            });
            Assert.Equal(3, attempts);
            Assert.False(Directory.Exists(sandbox));
        }
        finally { if (Directory.Exists(sandbox)) { Directory.Delete(sandbox, recursive: true); } }
    }

    [Fact]
    [Trait("Feature", "McpbPackaging")]
    public void RemoveStagingDirectory_WhenTimeoutExpires_ReportsTerminalState()
    {
        var sandbox = NewSandbox();
        try
        {
            var error = Assert.Throws<IOException>(() => PackageFiles.RemoveStaging(sandbox, TimeSpan.Zero, TimeSpan.Zero,
                _ => throw new UnauthorizedAccessException("locked by scanner")));
            Assert.True(Directory.Exists(sandbox));
            foreach (var text in new[] { "Failed to remove MCPB staging directory", sandbox, "after 1 attempt", "within 0 ms", "stale staging remains", "locked by scanner" })
            {
                Assert.Contains(text, error.Message, StringComparison.Ordinal);
            }
        }
        finally { Directory.Delete(sandbox, recursive: true); }
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [Trait("Feature", "McpbPackaging")]
    public void RemoveStagingDirectory_WhenTimeoutOrRetryIntervalIsNegative_UsesMatchingArgumentName(int timeoutMs, int retryMs)
    {
        var sandbox = NewSandbox();
        try
        {
            var timeout = TimeSpan.FromMilliseconds(timeoutMs);
            var retryInterval = TimeSpan.FromMilliseconds(retryMs);
            var error = Assert.ThrowsAny<ArgumentOutOfRangeException>(() => PackageFiles.RemoveStaging(sandbox, timeout, retryInterval));
            var expectedName = timeoutMs < 0 ? nameof(timeout) : nameof(retryInterval);
            Assert.Equal(expectedName, error.ParamName);
        }
        finally { if (Directory.Exists(sandbox)) { Directory.Delete(sandbox, recursive: true); } }
    }

    private static string NewSandbox() => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.TypedPackages.{Guid.NewGuid():N}")).FullName;
}

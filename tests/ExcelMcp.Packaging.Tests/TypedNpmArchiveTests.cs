using System.Formats.Tar;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "Packaging")]
public sealed class TypedNpmArchiveTests
{
    [Theory]
    [InlineData("../escape.txt", TarEntryType.RegularFile)]
    [InlineData("package/../../escape.txt", TarEntryType.RegularFile)]
    [InlineData("package/link", TarEntryType.SymbolicLink)]
    [InlineData("package/link", TarEntryType.HardLink)]
    public async Task ArchiveExtraction_RejectsTraversalAndLinks(string name, TarEntryType kind)
    {
        var sandbox = NewSandbox();
        try
        {
            var archive = Path.Combine(sandbox, "unsafe.tgz");
            WriteArchive(archive, [(name, kind, Encoding.UTF8.GetBytes("unsafe"))]);
            var result = await ExpandAsync(archive, Path.Combine(sandbox, "extracted"));
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains(kind == TarEntryType.RegularFile ? "Unsafe or duplicate" : "Unsupported", result.CombinedOutput, StringComparison.Ordinal);
            Assert.False(File.Exists(Path.Combine(sandbox, "escape.txt")));
        }
        finally { Directory.Delete(sandbox, true); }
    }

    [Fact]
    public async Task ArchiveExtraction_RejectsDuplicatePathsAndMalformedArchives()
    {
        var sandbox = NewSandbox();
        try
        {
            var archive = Path.Combine(sandbox, "duplicate.tgz");
            WriteArchive(archive, [
                ("package/payload", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("first")),
                ("package/PAYLOAD", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("second"))
            ]);
            var duplicate = await ExpandAsync(archive, Path.Combine(sandbox, "duplicates"));
            Assert.NotEqual(0, duplicate.ExitCode);
            Assert.Contains("Unsafe or duplicate", duplicate.CombinedOutput, StringComparison.Ordinal);
            File.WriteAllText(archive, "not a gzip archive");
            var malformed = await ExpandAsync(archive, Path.Combine(sandbox, "malformed"));
            Assert.NotEqual(0, malformed.ExitCode);
        }
        finally { Directory.Delete(sandbox, true); }
    }

    [Theory]
    [InlineData("valid")]
    [InlineData("name")]
    [InlineData("architecture")]
    [InlineData("machine")]
    [InlineData("license")]
    [InlineData("launcher")]
    [InlineData("dependency")]
    public async Task ArchiveVerification_RequiresMatchingMetadataArchitectureAndContents(string mutation)
    {
        var sandbox = NewSandbox();
        try
        {
            var runtime = JsonNode.Parse("""{"name":"@sbroenne/excelcli-win32-x64","version":"2.1.3","main":"excelcli.exe","os":["win32"],"cpu":["x64"]}""")!.AsObject();
            var launcher = JsonNode.Parse("""{"name":"@sbroenne/excelcli","version":"2.1.3","bin":{"excelcli":"bin/excelcli.js"},"optionalDependencies":{"@sbroenne/excelcli-win32-x64":"2.1.3","@sbroenne/excelcli-win32-arm64":"2.1.3"}}""")!.AsObject();
            var executable = new byte[128];
            BitConverter.GetBytes((ushort)0x5a4d).CopyTo(executable, 0);
            BitConverter.GetBytes(64).CopyTo(executable, 0x3c);
            BitConverter.GetBytes(0x00004550).CopyTo(executable, 64);
            BitConverter.GetBytes((ushort)(mutation == "machine" ? 0xaa64 : 0x8664)).CopyTo(executable, 68);
            if (mutation == "name") { runtime["name"] = "@sbroenne/unrelated"; }
            if (mutation == "architecture") { runtime["cpu"] = new JsonArray("arm64"); }
            if (mutation == "dependency") { launcher["optionalDependencies"]!["@sbroenne/excelcli-win32-arm64"] = "1.0.0"; }
            var runtimeEntries = new List<(string, TarEntryType, byte[])> {
                ("package/package.json", TarEntryType.RegularFile, Encoding.UTF8.GetBytes(runtime.ToJsonString())),
                ("package/excelcli.exe", TarEntryType.RegularFile, executable)
            };
            if (mutation != "license") { runtimeEntries.Add(("package/LICENSE", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("MIT"))); }
            var launcherEntries = new List<(string, TarEntryType, byte[])> {
                ("package/package.json", TarEntryType.RegularFile, Encoding.UTF8.GetBytes(launcher.ToJsonString())),
                ("package/LICENSE", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("MIT")),
                ("package/bin/excelcli.js", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("fixture"))
            };
            if (mutation != "launcher") { launcherEntries.Add(("package/lib/launcher.js", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("fixture"))); }
            var runtimePath = Path.Combine(sandbox, "runtime.tgz");
            var launcherPath = Path.Combine(sandbox, "launcher.tgz");
            WriteArchive(runtimePath, runtimeEntries);
            WriteArchive(launcherPath, launcherEntries);
            var result = await RunPowerShellFileAsync(Path.Combine(RepoRoot, "scripts", "Test-NpmPackages.ps1"),
                ["-Component", "Cli", "-Architecture", "x64", "-LauncherPackage", launcherPath, "-RuntimePackage", runtimePath, "-ArchiveOnly"]);
            if (mutation == "valid")
            {
                Assert.True(result.ExitCode == 0, result.CombinedOutput);
                Assert.Contains("Archive-only validation requested", result.CombinedOutput, StringComparison.Ordinal);
            }
            else
            {
                Assert.NotEqual(0, result.ExitCode);
                var expected = mutation switch
                {
                    "name" => "Unexpected npm package name",
                    "architecture" => "metadata does not match",
                    "machine" => "does not match x64",
                    "license" => "Missing license",
                    "launcher" => "Missing launcher file",
                    _ => "matching arm64 release version"
                };
                Assert.Contains(expected, result.CombinedOutput, StringComparison.Ordinal);
            }
        }
        finally { Directory.Delete(sandbox, true); }
    }

    private static string NewSandbox() =>
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.NpmArchive.{Guid.NewGuid():N}")).FullName;

    private static Task<ProcessResult> ExpandAsync(string archive, string destination) =>
        RunPowerShellCommandAsync($"""
            . '{Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")}'
            Expand-PackageTar '{archive}' '{destination}'
            """);

    private static void WriteArchive(string path, IEnumerable<(string Name, TarEntryType Kind, byte[] Contents)> entries)
    {
        using var file = File.Create(path);
        using var compressed = new GZipStream(file, CompressionLevel.Fastest);
        using var writer = new TarWriter(compressed);
        foreach (var (name, kind, contents) in entries)
        {
            var entry = new PaxTarEntry(kind, name);
            using var data = new MemoryStream(contents);
            if (kind == TarEntryType.RegularFile) { entry.DataStream = data; }
            else { entry.LinkName = "../escape.txt"; }
            writer.WriteEntry(entry);
        }
    }
}

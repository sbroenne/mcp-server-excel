using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("Category", "Integration")]
[Trait("Feature", "PreCommit")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookPackageAccessGuardTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string ScriptPath = Path.Combine(
        RepoRoot,
        "scripts",
        "check-workbook-package-access.ps1");

    [Fact]
    public async Task ProductionWorkbookPackageXmlAccess_IsRejected()
    {
        var result = await RunGuardAsync(
            ("src/Runtime/ZipXmlReader.cs", """
                using System.IO.Compression;
                using System.Xml.Linq;
                using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
                var document = XDocument.Load(archive.Entries[0].Open());
                """),
            ("src/Runtime/OpenXmlReader.cs", """
                using DocumentFormat.OpenXml.Packaging;
                using var document = SpreadsheetDocument.Open(path, false);
                """),
            ("src/Runtime/PartPathReader.cs", """
                const string WorkbookPartPath = "xl/workbook.xml";
                """));

        Assert.NotEqual(0, result.ExitCode);
        var output = result.Output.Replace('/', '\\');
        Assert.Contains(@"src\Runtime\ZipXmlReader.cs", output, StringComparison.Ordinal);
        Assert.Contains(@"src\Runtime\OpenXmlReader.cs", output, StringComparison.Ordinal);
        Assert.Contains(@"src\Runtime\PartPathReader.cs", output, StringComparison.Ordinal);
        Assert.Contains("Use Excel COM in production", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TestPackageAccessAndUnrelatedXmlOrZipProcessing_AreAllowed()
    {
        var result = await RunGuardAsync(
            ("tests/Fixtures/WorkbookPackageFixture.cs", """
                using System.IO.Compression;
                using System.Xml.Linq;
                const string WorkbookPart = "xl/workbook.xml";
                """),
            ("src/Runtime/XmlInputValidator.cs", """
                using System.Xml;
                using System.Xml.Linq;
                var document = XDocument.Parse(xml);
                """),
            ("src/Runtime/DownloadArchive.cs", """
                using System.IO.Compression;
                using var archive = ZipFile.OpenRead(path);
                """),
            ("src/Runtime/obj/GeneratedWorkbookReader.cs", """
                using System.IO.Compression;
                using System.Xml.Linq;
                const string WorkbookPart = "xl/workbook.xml";
                """));

        Assert.Equal(0, result.ExitCode);
        Assert.Contains(
            "No production Excel workbook package XML access found.",
            result.Output,
            StringComparison.Ordinal);
    }

    private static async Task<ScriptResult> RunGuardAsync(
        params (string RelativePath, string Content)[] files)
    {
        var sandbox = Path.Combine(
            Path.GetTempPath(),
            $"ExcelMcpPackageGuard-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path.Combine(sandbox, "src"));

        try
        {
            foreach (var (relativePath, content) in files)
            {
                var path = Path.Combine(
                    sandbox,
                    relativePath.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                await File.WriteAllTextAsync(path, content);
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = "pwsh",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            startInfo.Environment["EXCELMCP_BUILD_ROOT"] = RepoRoot;
            startInfo.Environment["EXCELMCP_BUILD_DLL"] = typeof(Sbroenne.ExcelMcp.Build.ValidationPolicy).Assembly.Location;
            startInfo.ArgumentList.Add("-NoProfile");
            startInfo.ArgumentList.Add("-File");
            startInfo.ArgumentList.Add(ScriptPath);
            startInfo.ArgumentList.Add("-RootPath");
            startInfo.ArgumentList.Add(sandbox);

            using var process = Process.Start(startInfo);
            Assert.NotNull(process);
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            await process.WaitForExitAsync(timeout.Token);

            return new ScriptResult(
                process.ExitCode,
                $"{await stdout}{Environment.NewLine}{await stderr}");
        }
        finally
        {
            Directory.Delete(sandbox, recursive: true);
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

        throw new DirectoryNotFoundException(
            "Could not locate repository root from test output directory.");
    }

    private sealed record ScriptResult(int ExitCode, string Output);
}

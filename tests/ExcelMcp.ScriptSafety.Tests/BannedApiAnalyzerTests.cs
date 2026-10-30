using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("Feature", "PreCommit")]
[Trait("RequiresExcel", "false")]
public sealed class BannedApiAnalyzerTests
{
    [Theory]
    [InlineData("System.Console", "BannedSymbols.Mcp.txt")]
    [InlineData("System.GC", "BannedSymbols.txt")]
    public void AllDangerousFrameworkOverloads_HaveAnExactBan(string typeName, string file)
    {
        var type = SafetyAnalyzerTests.CreateCompilation("").GetTypeByMetadataName(typeName);
        Assert.NotNull(type);
        var bans = File.ReadLines(Path.Combine(TypedValidationPolicyTests.Root, file))
            .Select(line => line.Split(';')[0]).ToHashSet(StringComparer.Ordinal);
        var methods = type.GetMembers().OfType<IMethodSymbol>().Where(method =>
            typeName == "System.Console" ? method.Name is "Write" or "WriteLine" : method.Name == "Collect").ToArray();
        Assert.NotEmpty(methods);
        Assert.All(methods, method =>
        {
            var id = method.GetDocumentationCommentId();
            Assert.NotNull(id);
            Assert.Contains(id, bans);
        });
    }

    [Fact]
    public async Task ActualPackage_TestCodeCannotSearchOrStopProcesses()
    {
        string[] calls =
        [
            "System.Diagnostics.Process.GetProcessesByName(\"probe\");",
            "System.Diagnostics.Process.GetProcessesByName(\"probe\", \".\");",
            "System.Diagnostics.Process.GetProcesses();",
            "System.Diagnostics.Process.GetProcesses(\".\");",
            "Sbroenne.ExcelMcp.ComInterop.Session.OwnedProcessGuard.TryTerminate(default, default, default, out _);",
            "Sbroenne.ExcelMcp.ComInterop.Session.ExcelBatch.TryTerminateOwnedProcess(default, default, default);"
        ];
        await AssertProbeRejectsEachLine(["BannedSymbols.TestProcesses.txt"], calls,
            """
            namespace Sbroenne.ExcelMcp.ComInterop.Session {
                public readonly record struct ExcelProcessIdentity(int ProcessId, long StartedAtUtcFileTime);
                public static class OwnedProcessGuard { public static bool TryTerminate(ExcelProcessIdentity identity, System.TimeSpan before, System.TimeSpan after, out bool terminated, System.TimeSpan? overall = null) { terminated = false; return false; } }
                public static class ExcelBatch { public static bool TryTerminateOwnedProcess(ExcelProcessIdentity identity, System.TimeSpan before, System.TimeSpan after) => false; }
            }
            """);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ActualPackage_EnforcesTheRepositoryBansAcrossOverloadsAndNamespaces(bool banned)
    {
        string[] calls = banned
            ? [
                "System.Console.Out.WriteLine(\"probe\");",
                "System.Console.Write(\"probe\");",
                "System.Console.WriteLine(\"probe\");",
                "System.Console.WriteLine(\"{0}\", 1);",
                "System.Console.SetOut(System.IO.TextWriter.Null);",
                "System.GC.Collect();",
                "System.GC.Collect(0);",
                "System.GC.Collect(0, System.GCCollectionMode.Forced);",
                "System.GC.Collect(0, System.GCCollectionMode.Forced, true);",
                "System.GC.Collect(0, System.GCCollectionMode.Forced, true, true);",
                "System.GC.WaitForPendingFinalizers();",
                "System.Runtime.InteropServices.Marshal.CleanupUnusedObjectsInCurrentContext();",
                "System.Diagnostics.Process.GetProcessesByName(\"probe\");",
                "System.Diagnostics.Process.GetProcessesByName(\"probe\", \".\");",
                "System.Runtime.InteropServices.Marshal.ReleaseComObject(new object());",
                "System.Runtime.InteropServices.Marshal.FinalReleaseComObject(new object());",
                "_ = new DocumentFormat.OpenXml.Document();",
                "_ = new DocumentFormat.OpenXml.Packaging.Document();",
                "_ = new System.IO.Packaging.Package();",
                "book.RefreshAll();"
            ]
            : [
                "System.Console.Error.WriteLine(\"probe\");",
                "using var process = System.Diagnostics.Process.GetProcessById(System.Environment.ProcessId);",
                "_ = System.GC.GetTotalMemory(false);"
            ];
        var output = await BuildProbe(["BannedSymbols.txt", "BannedSymbols.Mcp.txt"], calls,
            """
            namespace DocumentFormat.OpenXml { public class Document {} }
            namespace DocumentFormat.OpenXml.Packaging { public class Document {} }
            namespace System.IO.Packaging { public class Package {} }
            """);
        if (!banned)
        {
            Assert.True(output.ExitCode == 0, output.Text);
            return;
        }
        AssertEachLineRejected(output, calls.Length);
    }

    private static async Task AssertProbeRejectsEachLine(string[] banFiles, string[] calls, string stubs) =>
        AssertEachLineRejected(await BuildProbe(banFiles, calls, stubs), calls.Length);

    private static void AssertEachLineRejected((int ExitCode, string Text) output, int lineCount)
    {
        Assert.NotEqual(0, output.ExitCode);
        var rejectedLines = Regex.Matches(output.Text, @"Probe\.cs\((\d+),\d+\): error RS0030")
            .Select(match => int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture))
            .Distinct().Order().ToArray();
        Assert.True(Enumerable.Range(2, lineCount).SequenceEqual(rejectedLines), output.Text);
    }

    private static async Task<(int ExitCode, string Text)> BuildProbe(string[] banFiles, string[] calls, string stubs)
    {
        var root = TypedValidationPolicyTests.Root;
        var directory = Path.Combine(Path.GetTempPath(), $"ExcelMcp.BannedApi.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var version = XDocument.Load(Path.Combine(root, "Directory.Packages.props"))
                .Descendants("PackageVersion")
                .Single(element => (string?)element.Attribute("Include") == "Microsoft.CodeAnalysis.BannedApiAnalyzers")
                .Attribute("Version")?.Value;
            Assert.NotNull(version);
            Assert.False(string.IsNullOrWhiteSpace(version));
            var project = Path.Combine(directory, "CompilerBanProbe.csproj");
            new XDocument(new XElement("Project", new XAttribute("Sdk", "Microsoft.NET.Sdk"),
                new XElement("PropertyGroup",
                    new XElement("TargetFramework", "net10.0-windows"),
                    new XElement("EnableWindowsTargeting", "true"),
                    new XElement("TreatWarningsAsErrors", "true"),
                    new XElement("NuGetAudit", "false")),
                new XElement("ItemGroup",
                    new XElement("PackageReference",
                        new XAttribute("Include", "Microsoft.CodeAnalysis.BannedApiAnalyzers"),
                        new XAttribute("Version", version)),
                    banFiles.Select(file => new XElement("AdditionalFiles", new XAttribute("Include", Path.Combine(root, file)))))))
                .Save(project);
            await File.WriteAllTextAsync(Path.Combine(directory, "Probe.cs"),
                "class Probe { static void M(Microsoft.Office.Interop.Excel._Workbook book) {\n" +
                string.Join('\n', calls) + "\n} }\n" +
                "namespace Microsoft.Office.Interop.Excel { public interface _Workbook { void RefreshAll(); } }\n" +
                stubs);
            var result = await new ScriptTestProcess(root).RunAsync("dotnet",
                ["build", project, "-c", "Release", "--disable-build-servers", "--nologo", "--verbosity", "minimal"],
                TimeSpan.FromMinutes(3));
            var output = result.Output + result.Error;
            Assert.DoesNotContain("error CS", output, StringComparison.Ordinal);
            return (result.ExitCode, output);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }
}

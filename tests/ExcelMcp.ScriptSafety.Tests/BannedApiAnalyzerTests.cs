using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Sbroenne.ExcelMcp.Build;
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

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ActualPackage_EnforcesTheRepositoryBansAcrossOverloadsAndNamespaces(bool banned)
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
                    new XElement("AdditionalFiles", new XAttribute("Include", Path.Combine(root, "BannedSymbols.txt"))),
                    new XElement("AdditionalFiles", new XAttribute("Include", Path.Combine(root, "BannedSymbols.Mcp.txt"))))))
                .Save(project);
            string[] calls = banned
                ? [
                    "System.Console.Out.WriteLine(\"probe\");",
                    "System.Console.Write(\"probe\");",
                    "System.Console.WriteLine(\"probe\");",
                    "System.Console.WriteLine(\"{0}\", 1);",
                    "System.Console.SetOut(System.IO.TextWriter.Null);",
                    "System.GC.Collect();",
                    "System.GC.Collect(0);",
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
            await File.WriteAllTextAsync(Path.Combine(directory, "Probe.cs"),
                "class Probe { static void M(Microsoft.Office.Interop.Excel._Workbook book) {\n" +
                string.Join('\n', calls) + "\n} }\n" +
                """
                namespace DocumentFormat.OpenXml { public class Document {} }
                namespace DocumentFormat.OpenXml.Packaging { public class Document {} }
                namespace System.IO.Packaging { public class Package {} }
                namespace Microsoft.Office.Interop.Excel { public interface _Workbook { void RefreshAll(); } }
                """);
            var result = await new ProcessRunner(root).RunAsync("dotnet",
                ["build", project, "-c", "Release", "--disable-build-servers", "--nologo", "--verbosity", "minimal"],
                TimeSpan.FromMinutes(3));
            var output = result.Output + result.Error;
            Assert.DoesNotContain("error CS", output, StringComparison.Ordinal);
            if (!banned)
            {
                Assert.True(result.ExitCode == 0, output);
                return;
            }
            Assert.NotEqual(0, result.ExitCode);
            var rejectedLines = Regex.Matches(output, @"Probe\.cs\((\d+),\d+\): error RS0030")
                .Select(match => int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture))
                .Distinct().Order().ToArray();
            Assert.True(Enumerable.Range(2, calls.Length).SequenceEqual(rejectedLines), output);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }
}

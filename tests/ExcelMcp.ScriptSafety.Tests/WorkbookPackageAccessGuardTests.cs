using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("Feature", "PreCommit")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookPackageAccessGuardTests
{
    [Theory]
    [InlineData("Debug")]
    [InlineData("Release")]
    public async Task OrdinaryDirectAndIncrementalBuilds_EnforceAllFourPolicies(string configuration)
    {
        using var fixture = new CompilerBuildFixture();
        Assert.False(File.Exists(fixture.AnalyzerAssembly(configuration)));
        await fixture.WriteSourceAsync("namespace Fixture; public sealed class Legal { public static int Read(int value) => value; }");
        var baseline = await fixture.BuildAsync(configuration);
        Assert.True(baseline.ExitCode == 0, baseline.Output + baseline.Error);
        Assert.DoesNotContain("CS8032", baseline.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("CS9057", baseline.Output, StringComparison.Ordinal);
        Assert.True(File.Exists(fixture.AnalyzerAssembly(configuration)));
        Assert.False(File.Exists(Path.Combine(fixture.BuildArtifacts, "bin", "Fixture", configuration.ToLowerInvariant(), "ExcelMcp.Analyzers.dll")));

        var mutations = new (string Source, string Id)[]
        {
            ("internal class Bad { dynamic? field; }", "EXCEL001"),
            ("""
                internal class Bad {
                    // PIA gap: Office.Core IDispatch numeric property.
                    int Read(dynamic value) => (int)value;
                }
                """, "EXCEL002"),
            ("internal class Bad { int Read(Microsoft.Office.Interop.Excel.Range range) => range.Rows.Count; }", "EXCEL003"),
            ("""
                namespace Sbroenne.ExcelMcp.Core.Models {
                    internal class ResultBase { public bool Success { get; set; } public string? ErrorMessage { get; set; } }
                }
                internal class Bad {
                    object Read() => new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = "failure" };
                }
                """, "EXCEL004"),
            ("""
                namespace Sbroenne.ExcelMcp.Core.Models {
                    internal class ResultBase { public bool Success { get; set; } public string? ErrorMessage { get; set; } }
                }
                internal class Bad {
                    static object Read() {
                        var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                        var alias = result;
                        alias.ErrorMessage = "failure";
                        result.Success = true;
                        return result;
                    }
                }
                """, "EXCEL004"),
            ("internal class Bad { string path = \"xl/workbook.xml\"; }", "EXCEL006"),
            ("""
                namespace DocumentFormat.OpenXml { internal class Workbook {} }
                internal class Bad { object Read() => new DocumentFormat.OpenXml.Workbook(); }
                """, "RS0030"),
            ("""
                internal class Bad {
                    void Read(System.IO.Compression.ZipArchive archive) {
                        var document = System.Xml.Linq.XDocument.Load(archive.GetEntry("xl/" + "workbook.xml")!.Open());
                    }
                }
                """, "EXCEL006")
        };
        foreach (var (source, id) in mutations)
        {
            await fixture.WriteSourceAsync(source);
            var result = await fixture.BuildAsync(configuration, noRestore: true);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("error " + id, result.Output, StringComparison.Ordinal);
            Assert.Contains("Cases.cs(", result.Output, StringComparison.Ordinal);
            Assert.DoesNotContain("AD0001", result.Output, StringComparison.Ordinal);
            await fixture.WriteSourceAsync("namespace Fixture; public sealed class Legal { public static int Read(int value) => value; }");
            var repaired = await fixture.BuildAsync(configuration, noRestore: true);
            Assert.True(repaired.ExitCode == 0, repaired.Output + repaired.Error);
        }
        var incremental = await fixture.BuildAsync(configuration, noRestore: true);
        Assert.True(incremental.ExitCode == 0, incremental.Output + incremental.Error);
    }

    [Fact]
    public async Task EvaluatedImportedAndTransitiveDependencies_AreRejectedEvenWithoutCompilation()
    {
        using var fixture = new CompilerBuildFixture();
        await fixture.WriteSourceAsync("namespace Fixture; public sealed class Legal {}");
        var baseline = await fixture.BuildAsync("Release");
        Assert.True(baseline.ExitCode == 0, baseline.Output + baseline.Error);
        foreach (var item in new[]
        {
            """<PackageReference Include="System.IO.Packaging" Version="10.0.0" />""",
            """<Reference Include="DocumentFormat.OpenXml" />""",
            """<ReferencePath Include="indirect.dll"><NuGetPackageId>DocumentFormat.OpenXml.Framework</NuGetPackageId></ReferencePath>"""
        })
        {
            await File.WriteAllTextAsync(Path.Combine(fixture.Root, "Dependencies.targets"),
                "<Project><ItemGroup>" + item + "</ItemGroup></Project>");
            var result = await fixture.RunAsync("msbuild", fixture.Project, "-t:RejectWorkbookPackageDependencies", "-v:minimal");
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("error EXCEL005", result.Output, StringComparison.Ordinal);
            Assert.Contains("Use Excel COM in production", result.Output, StringComparison.Ordinal);
        }
        await File.WriteAllTextAsync(Path.Combine(fixture.Root, "Dependencies.targets"),
            "<Project><PropertyGroup><WorkbookPart>xl/workbook.xml</WorkbookPart></PropertyGroup></Project>");
        var importedLiteral = await fixture.BuildAsync("Release", noRestore: true);
        Assert.NotEqual(0, importedLiteral.ExitCode);
        Assert.Contains("error EXCEL006", importedLiteral.Output, StringComparison.Ordinal);
        Assert.Contains("Dependencies.targets(", importedLiteral.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TestOnlyOoxmlAndSeparateZipXml_RemainLegal()
    {
        using var fixture = new CompilerBuildFixture(production: false);
        await fixture.WriteSourceAsync("""
            namespace DocumentFormat.OpenXml { public sealed class Workbook {} }
            namespace TestFixtures { public sealed class Fixture {
                public static object Read() => new DocumentFormat.OpenXml.Workbook();
                public static string Part() => "xl/workbook.xml";
            } }
            """);
        var tests = await fixture.BuildAsync("Release");
        Assert.True(tests.ExitCode == 0, tests.Output + tests.Error);

        using var production = new CompilerBuildFixture();
        await production.WriteSourceAsync("""
            namespace Fixture;
            public sealed class Legal {
                public static void Read(System.IO.Stream stream, string xml) {
                    using var archive = new System.IO.Compression.ZipArchive(stream);
                    var document = System.Xml.Linq.XDocument.Parse(xml);
                }
            }
            """);
        var separate = await production.BuildAsync("Release");
        Assert.True(separate.ExitCode == 0, separate.Output + separate.Error);
    }

    [Theory]
    [InlineData("result.ErrorMessage = null;", true)]
    [InlineData("result.Success = false;", true)]
    [InlineData("if (condition) result.ErrorMessage = null; else result.Success = false;", true)]
    [InlineData("if (condition) result.ErrorMessage = \"first\"; else result.ErrorMessage = \"second\";", false)]
    [InlineData("if (condition) { result.ErrorMessage = \"failure\"; return result; } result.ErrorMessage = null;", false)]
    public async Task OrdinaryBuild_ChecksTerminalResultState(string change, bool legal)
    {
        using var fixture = new CompilerBuildFixture();
        await fixture.WriteSourceAsync($$"""
            namespace Sbroenne.ExcelMcp.Core.Models {
                public class ResultBase { public bool Success { get; set; } public string? ErrorMessage { get; set; } }
            }
            namespace Fixture { public class Cases {
                public static object Read(bool condition) {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = {{(legal ? "\"failure\"" : "null")}} };
                    {{change}}
                    return result;
                } }
            }
            """);
        var build = await fixture.BuildAsync("Release");
        if (legal) { Assert.True(build.ExitCode == 0, build.Output + build.Error); }
        else
        {
            Assert.NotEqual(0, build.ExitCode);
            Assert.Contains("error EXCEL004", build.Output, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task OrdinaryBuild_AllowsSettingsZipAndGzipXmlButRejectsWorkbookHelper()
    {
        using var fixture = new CompilerBuildFixture();
        await fixture.WriteSourceAsync("""
            namespace Fixture;
            public class Cases {
                public static void Read(System.IO.Stream stream, System.IO.Compression.ZipArchive archive) {
                    using var settings = archive.GetEntry("settings.xml")!.Open();
                    var document = System.Xml.Linq.XDocument.Load(settings);
                    using var gzip = new System.IO.Compression.GZipStream(stream, System.IO.Compression.CompressionMode.Decompress);
                    var compressed = System.Xml.Linq.XDocument.Load(gzip);
                }
            }
            """);
        var legal = await fixture.BuildAsync("Release");
        Assert.True(legal.ExitCode == 0, legal.Output + legal.Error);
        await fixture.WriteSourceAsync("""
            namespace Fixture;
            public class Cases {
                private static System.IO.Stream Entry(System.IO.Compression.ZipArchive archive) => archive.Entries[0].Open();
                private static void Read(System.IO.Stream stream) {
                    var document = System.Xml.Linq.XDocument.Load(stream);
                }
                public static void Workbook() {
                    using var archive = System.IO.Compression.ZipFile.OpenRead("report.xlsx");
                    Read(Entry(archive));
                }
            }
            """);
        var illegal = await fixture.BuildAsync("Release", noRestore: true);
        Assert.NotEqual(0, illegal.ExitCode);
        Assert.Contains("error EXCEL006", illegal.Output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task CompatibilityBuildCommand_RejectsUnknownRulesAndEmptyProjectDiscovery()
    {
        using var fixture = new CompilerBuildFixture();
        var runner = new ProcessRunner(fixture.Root);
        var guards = new SourceGuards(fixture.Root);
        await Assert.ThrowsAsync<ArgumentException>(() => guards.CheckAsync("typo", null, runner));
        await Assert.ThrowsAsync<ArgumentException>(() => guards.CheckAsync("com-leaks", [], runner));
        await Assert.ThrowsAsync<DirectoryNotFoundException>(() => guards.CheckAsync("dynamic-casts", null, runner));
        Directory.CreateDirectory(Path.Combine(fixture.Root, "src"));
        await Assert.ThrowsAsync<InvalidOperationException>(() => guards.CheckAsync("dynamic-casts", null, runner));
    }

    [Fact]
    public async Task CompatibilityBuildCommand_RejectsProjectsWithoutProductionPolicy()
    {
        using var fixture = new CompilerBuildFixture();
        var source = Path.Combine(fixture.Root, "src", "Runtime");
        Directory.CreateDirectory(source);
        await File.WriteAllTextAsync(Path.Combine(source, "Runtime.csproj"),
            """<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><TargetFramework>net10.0</TargetFramework></PropertyGroup></Project>""");
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            new SourceGuards(fixture.Root).CheckAsync("dynamic-casts", null, new ProcessRunner(fixture.Root)));
        Assert.Contains("No source check was performed", error.Message, StringComparison.Ordinal);
    }
}

internal sealed class CompilerBuildFixture : IDisposable
{
    public string Root { get; }
    public string Project { get; }
    public string BuildArtifacts => Path.Combine(Root, "out");

    public string AnalyzerAssembly(string configuration) =>
        Path.Combine(BuildArtifacts, "bin", "ExcelMcp.Analyzers", configuration.ToLowerInvariant(), "ExcelMcp.Analyzers.dll");

    public CompilerBuildFixture(bool production = true)
    {
        var repo = TypedSourceGuardTests.FindRoot();
        Root = Path.Combine(repo, "artifacts", "guard-builds", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Root);
        Project = Path.Combine(Root, "Fixture.csproj");
        File.WriteAllText(Project, $$"""
            <Project Sdk="Microsoft.NET.Sdk">
              <PropertyGroup>
                <TargetFramework>net10.0</TargetFramework>
                <ExcelMcpProductionSafety>{{production.ToString().ToLowerInvariant()}}</ExcelMcpProductionSafety>
                <IsTestProject>{{(!production).ToString().ToLowerInvariant()}}</IsTestProject>
                <IsPackable>false</IsPackable>
                <NoWarn>$(NoWarn);CS1591;CS0169;CS0414</NoWarn>
              </PropertyGroup>
              <ItemGroup>
                <PackageReference Include="Microsoft.Office.Interop.Excel" />
              </ItemGroup>
              <Import Project="Dependencies.targets" />
            </Project>
            """);
        File.WriteAllText(Path.Combine(Root, "Dependencies.targets"), "<Project />");
        File.WriteAllText(Path.Combine(Root, ".editorconfig"), "root = true\n[*.cs]\nexcelmcp_typed_excel = true\n");
    }

    public Task WriteSourceAsync(string source) => File.WriteAllTextAsync(Path.Combine(Root, "Cases.cs"), source);

    public Task<ProcessResult> BuildAsync(string configuration, bool noRestore = false)
    {
        var arguments = new List<string> { "build", Project, "-c", configuration, "--disable-build-servers", "--verbosity", "minimal" };
        arguments.Add("-p:UseArtifactsOutput=true");
        arguments.Add("-p:ArtifactsPath=" + BuildArtifacts);
        if (noRestore) { arguments.Add("--no-restore"); }
        return RunAsync(arguments.ToArray());
    }

    public Task<ProcessResult> RunAsync(params string[] arguments) =>
        new ProcessRunner(Root).RunAsync("dotnet", arguments, TimeSpan.FromMinutes(2));

    public void Dispose() => Directory.Delete(Root, recursive: true);
}

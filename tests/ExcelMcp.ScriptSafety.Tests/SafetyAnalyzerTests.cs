using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;
using Sbroenne.ExcelMcp.Analyzers;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("Feature", "PreCommit")]
[Trait("RequiresExcel", "false")]
public sealed class SafetyAnalyzerTests
{
    private static readonly ImmutableArray<MetadataReference> References =
        ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator)
        .Select(path => MetadataReference.CreateFromFile(path)).ToImmutableArray<MetadataReference>();

    [Theory]
    [InlineData("catch {}", "EXL001")]
    [InlineData("catch { /* ignore */ }", "EXL001")]
    [InlineData("catch (COMException) { return false; }", "EXL002")]
    [InlineData("catch (Exception) { return false; }", "EXL002")]
    [InlineData("catch (Exception ex) when (ex is COMException) { return false; }", "EXL002")]
    [InlineData("catch (COMException ex) when (ex.HResult == GenericExcelFailure) { return false; }", "EXL002")]
    public async Task Catch_RejectsHiddenFailures(string handler, string id)
    {
        var diagnostics = await AnalyzeAsync($$"""
            using System;
            using System.Runtime.InteropServices;
            class C {
                const int GenericExcelFailure = unchecked((int)0x800A03EC);
                bool M() { try { return true; } {{handler}} return false; }
            }
            """);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Id == id);
    }

    [Theory]
    [InlineData("catch { throw; }")]
    [InlineData("catch (COMException ex) { throw new InvalidOperationException(\"Failed\", ex); }")]
    [InlineData("catch (COMException ex) when (ex.HResult == MemberNotFound) { return false; }")]
    [InlineData("catch (COMException ex) when (ex.HResult == MemberNotFound) { }")]
    [InlineData("catch (COMException ex) when (IsUnavailable(ex)) { return false; }")]
    public async Task Catch_AllowsPropagationAndSpecificExpectedFailure(string handler)
    {
        var diagnostics = await AnalyzeAsync($$"""
            using System;
            using System.Runtime.InteropServices;
            class C {
                const int MemberNotFound = unchecked((int)0x80020003);
                static bool IsUnavailable(COMException ex) => ex.HResult == MemberNotFound;
                bool M() { try { return true; } {{handler}} return false; }
            }
            """);
        Assert.Empty(diagnostics);
    }

    [Theory]
    [InlineData("catch (COMException ex) when (IsCom(ex)) { return false; }")]
    [InlineData("catch (COMException ex) when (IsUnavailable(ex) || DateTime.Now.Ticks > 0) { return false; }")]
    [InlineData("catch (COMException ex) when (IsUnavailable(ex) || IsCom(ex)) { return false; }")]
    [InlineData("catch (COMException ex) when (IsUnavailable(new COMException(\"\", MemberNotFound))) { return false; }")]
    public async Task Catch_FilterCannotDisguiseAnUnrestrictedFallback(string handler)
    {
        var diagnostics = await AnalyzeAsync($$"""
            using System;
            using System.Runtime.InteropServices;
            class C {
                const int MemberNotFound = unchecked((int)0x80020003);
                static bool IsUnavailable(COMException ex) => ex.HResult == MemberNotFound;
                static bool IsCom(Exception ex) => ex is COMException;
                bool M() { try { return true; } {{handler}} }
            }
            """);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Id == "EXL002");
    }

    [Fact]
    public async Task Catch_TracksAnExceptionNarrowedByATypePattern()
    {
        var diagnostics = await AnalyzeAsync("""
            using System;
            using System.Runtime.InteropServices;
            class C {
                const int MemberNotFound = unchecked((int)0x80020003);
                static bool IsUnavailable(Exception ex) =>
                    ex is COMException native && IsNativeUnavailable(native);
                static bool IsNativeUnavailable(COMException ex) => ex.HResult == MemberNotFound;
                bool M() {
                    try { return true; }
                    catch (Exception ex) when (IsUnavailable(ex)) { return false; }
                }
            }
            """);
        Assert.Empty(diagnostics);
    }

    [Fact]
    public async Task ComEnums_AreValuesNotAcquiredObjects()
    {
        var diagnostics = await AnalyzeAsync("""
            namespace Microsoft.Office.Interop.Excel {
                enum XlType { Default }
                interface Range { XlType Type { get; } }
            }
            class C { int M(Microsoft.Office.Interop.Excel.Range range) {
                return System.Convert.ToInt32(range.Type);
            } }
            """, new ComAccessAnalyzer());
        Assert.Empty(diagnostics);
    }

    [Theory]
    [InlineData("#pragma warning disable CS8602", true)]
    [InlineData("#pragma warning disable CS8602 // Excel returned a documented nullable reference.", false)]
    public async Task Suppression_NeedsReason(string directive, bool expected)
    {
        var diagnostics = await AnalyzeAsync(directive + "\nclass C {}");
        Assert.Equal(expected, diagnostics.Any(diagnostic => diagnostic.Id == "EXL006"));
    }

    [Theory]
    [InlineData("int x = (int)value;", true)]
    [InlineData("int x = Convert.ToInt32(value);", false)]
    public async Task DynamicNumber_UsesConversion(string body, bool expected)
    {
        var diagnostics = await AnalyzeAsync("using System; class C { void M(dynamic value) { " + body + " } }");
        Assert.Equal(expected, diagnostics.Any(diagnostic => diagnostic.Id == "EXL010"));
    }

    [Theory]
    [InlineData("return range.Rows.Count;", true)]
    [InlineData("var rows = range.Rows; return rows.Count;", false)]
    [InlineData("return range.Count;", false)]
    public async Task ComAccess_UsesSymbolTypeNotSpelling(string body, bool expected)
    {
        var diagnostics = await AnalyzeAsync("""
            namespace Microsoft.Office.Interop.Excel {
                interface Range { Range Rows { get; } int Count { get; } }
            }
            class C {
                int M(Microsoft.Office.Interop.Excel.Range range) {
            """ + body + " } }", new ComAccessAnalyzer());
        Assert.Equal(expected, diagnostics.Any(diagnostic => diagnostic.Id == "EXL004"));
        var managed = await AnalyzeAsync("""
            class ManagedRows { public ManagedRows Rows => this; public int Count => 1; }
            class C { int M(ManagedRows range) {
            """ + body + " } }", new ComAccessAnalyzer());
        Assert.Empty(managed);
    }

    [Theory]
    [InlineData("names.Add();", true)]
    [InlineData("var result = names.Add();", false)]
    public async Task ComResult_MustNotBeDiscarded(string body, bool expected)
    {
        var diagnostics = await AnalyzeAsync("""
            namespace Microsoft.Office.Interop.Excel {
                interface Name {}
                interface Names { Name Add(); }
            }
            class C { void M(Microsoft.Office.Interop.Excel.Names names) {
            """ + body + " } }", new ComAccessAnalyzer());
        Assert.Equal(expected, diagnostics.Any(diagnostic => diagnostic.Id == "EXL007"));
    }

    [Theory]
    [InlineData("xl/workbook.xml", true)]
    [InlineData("[Content_Types].xml", true)]
    [InlineData("xl/worksheets/sheet1.xml", true)]
    [InlineData("input.xml", false)]
    public async Task WorkbookParts_OnlyFixtureCodeMayAccessThem(string value, bool expected)
    {
        var diagnostics = await AnalyzeAsync("class C { const string Part = \"" + value + "\"; }", new ComAccessAnalyzer());
        Assert.Equal(expected, diagnostics.Any(diagnostic => diagnostic.Id == "EXL009"));
    }

    [Theory]
    [InlineData("src\\ExcelMcp.Core\\Commands\\Settings\\NativeQueryApprovalSettings.cs", false)]
    [InlineData("src\\ExcelMcp.Core\\Commands\\Settings\\OtherSettings.cs", true)]
    public async Task ArchiveAccess_IsAllowedOnlyForNativeQuerySettings(string sourcePath, bool expected)
    {
        const string source = """
            using System.IO;
            using System.IO.Compression;
            class C {
                void M(Stream stream, string path) {
                    using var archive = new ZipArchive(stream);
                    using var other = ZipFile.OpenRead(path);
                }
            }
            """;
        var diagnostics = await AnalyzeAsync(source, new ComAccessAnalyzer(), path: sourcePath);
        Assert.Equal(expected, diagnostics.Any(diagnostic => diagnostic.Id == "EXL009"));
    }

    [Theory]
    [InlineData("", "EXL020")]
    [InlineData("[Trait(\"RequiresExcel\", \"sometimes\")]", "EXL020")]
    [InlineData("[Trait(\"RequiresExcel\", \"true\")]", "EXL021")]
    [InlineData("[Trait(\"RequiresExcel\", \"false\")]", null)]
    [InlineData("[Trait(\"RequiresExcel\", \"true\"), Collection(\"Excel\")]", null)]
    public async Task TestClassification_FailsBeforeStartingExcel(string attributes, string? expected)
    {
        var diagnostics = await AnalyzeAsync("""
            using System;
            using Xunit;
            namespace Xunit {
                class FactAttribute : Attribute {}
                [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple=true)]
                class TraitAttribute(string key, string value) : Attribute {}
                class CollectionAttribute(string name) : Attribute {}
                class CollectionDefinitionAttribute(string name) : Attribute {
                    public bool DisableParallelization { get; set; }
                }
            }
            [CollectionDefinition("Excel", DisableParallelization=true)] class ExcelCollection {}
            """ + attributes + " class C { [Fact] public void M() {} }", new TestClassificationAnalyzer());
        if (expected == null) { Assert.Empty(diagnostics); }
        else { Assert.Equal(expected, Assert.Single(diagnostics).Id); }
    }

    internal static async Task<ImmutableArray<Diagnostic>> AnalyzeAsync(string source, DiagnosticAnalyzer? analyzer = null,
        string assemblyName = "Fixture", ImmutableArray<AdditionalText> documents = default,
        string path = "src\\ExcelMcp.Core\\Fixture.cs")
    {
        var compilation = CreateCompilation(source, assemblyName, path);
        var errors = compilation.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error);
        Assert.Empty(errors);
        return await compilation.WithAnalyzers([analyzer ?? new SafetyAnalyzer()],
            new AnalyzerOptions(documents.IsDefault ? [] : documents)).GetAnalyzerDiagnosticsAsync();
    }

    internal static CSharpCompilation CreateCompilation(
        string source,
        string assemblyName = "Fixture",
        string path = "src\\ExcelMcp.Core\\Fixture.cs") =>
        CSharpCompilation.Create(assemblyName,
            [CSharpSyntaxTree.ParseText(source, path: path)], References,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
}

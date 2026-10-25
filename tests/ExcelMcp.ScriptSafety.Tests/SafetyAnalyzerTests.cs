using System.Collections.Immutable;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
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

    private static readonly MetadataReference ExcelReference = MetadataReference.CreateFromFile(
        Path.Combine(AppContext.BaseDirectory, "AnalyzerFixtures", "Microsoft.Office.Interop.Excel.dll"),
        new MetadataReferenceProperties(embedInteropTypes: true));

    private static readonly MetadataReference OfficeReference = MetadataReference.CreateFromFile(
        Path.Combine(AppContext.BaseDirectory, "AnalyzerFixtures", "office.dll"),
        new MetadataReferenceProperties(embedInteropTypes: true));

    private static readonly MetadataReference VbeReference = MetadataReference.CreateFromFile(
        Path.Combine(AppContext.BaseDirectory, "AnalyzerFixtures", "Microsoft.Vbe.Interop.dll"),
        new MetadataReferenceProperties(embedInteropTypes: true));

    private static readonly MetadataReference AdoReference = MetadataReference.CreateFromFile(
        Path.Combine(AppContext.BaseDirectory, "AnalyzerFixtures", "ADODB.dll"),
        new MetadataReferenceProperties(embedInteropTypes: true));

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
    [InlineData("Fixture")]
    [InlineData("Sbroenne.ExcelMcp.Service.Tests")]
    public async Task DynamicAccess_IsRejectedInProductionAndTestProjects(string assemblyName)
    {
        const string source = """
            class C {
                void M(dynamic value) {
                    value.Refresh();
                }
            }
            """;
        var diagnostics = await AnalyzeAsync(source, assemblyName: assemblyName);

        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Id == "EXL011" &&
            source.Substring(diagnostic.Location.SourceSpan.Start, diagnostic.Location.SourceSpan.Length)
                .Contains("value.Refresh", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("object formula = (object)range.Formula; object formula2 = (object)range.Formula2; object value = range.Value2; return formula;", false)]
    [InlineData("object value = range.Value2.Missing; return value;", true)]
    [InlineData("range.Value2.Refresh(); return null;", true)]
    [InlineData("var value = range.Value2; return (object)value;", true)]
    public async Task DynamicAccess_DistinguishesTypedExcelVariantReadsFromLateBinding(string body, bool expected)
    {
        var compilation = CreateCompilation($$"""
            using Excel = Microsoft.Office.Interop.Excel;
            class C {
                object M(Excel.Range range) {
                    {{body}}
                }
            }
            """).AddReferences(ExcelReference);
        Assert.Empty(compilation.GetDiagnostics().Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        var diagnostics = await compilation.WithAnalyzers([new SafetyAnalyzer()]).GetAnalyzerDiagnosticsAsync();

        Assert.Equal(expected, diagnostics.Any(diagnostic => diagnostic.Id == "EXL011"));
        if (!expected)
            Assert.Empty(diagnostics);
    }

    [Theory]
    [InlineData("SourceName")]
    [InlineData("RecordCount")]
    public void TypedMetadata_UnsupportedWorksheetPropertiesFailDuringCompilation(string property)
    {
        var compilation = CreateCompilation($$"""
            using Excel = Microsoft.Office.Interop.Excel;
            class C { object Read(Excel.Worksheet sheet) => sheet.{{property}}; }
            """).AddReferences(ExcelReference);

        var error = Assert.Single(compilation.GetDiagnostics()
            .Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        Assert.Equal("CS1061", error.Id);
        Assert.Contains(property, error.GetMessage(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("ModelTableColumn", "IsCalculatedColumn")]
    [InlineData("WorkbookConnection", "WebConnection")]
    public void TypedMetadata_RemainingCapabilityProbesHaveNoSupportedMember(string type, string property)
    {
        var compilation = CreateCompilation($$"""
            using Excel = Microsoft.Office.Interop.Excel;
            class C { object Read(Excel.{{type}} value) => value.{{property}}; }
            """).AddReferences(ExcelReference, OfficeReference, VbeReference);

        var error = Assert.Single(compilation.GetDiagnostics()
            .Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        Assert.Equal("CS1061", error.Id);
        Assert.Contains(property, error.GetMessage(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Excel.Workbook book", "book.VBProject", "Microsoft.Vbe.Interop")]
    [InlineData("Excel.Shape shape", "shape.Type", "office")]
    public void TypedMetadata_ExternalObjectModelsRequireTheirOwnTypeDefinitions(
        string parameter, string expression, string assemblyName)
    {
        var compilation = CreateCompilation($$"""
            using Excel = Microsoft.Office.Interop.Excel;
            class C { object Read({{parameter}}) => {{expression}}; }
            """).AddReferences(ExcelReference);

        var error = Assert.Single(compilation.GetDiagnostics()
            .Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        Assert.Equal("CS0012", error.Id);
        Assert.Contains(assemblyName, error.GetMessage(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    [Fact]
    public async Task TypedOfficeDefinitions_CompileAndEmbedWithoutRuntimeAssemblyReferences()
    {
        var compilation = CreateCompilation("""
            using Excel = Microsoft.Office.Interop.Excel;
            using Office = Microsoft.Office.Core;
            using Vbe = Microsoft.Vbe.Interop;
            class C {
                void Secure(Excel.Application app) =>
                    app.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                Office.MsoShapeType ReadShapeType(Excel.Shape shape) => shape.Type;
                string ReadModule(Vbe.VBComponent component) {
                    Vbe.CodeModule code = component.CodeModule;
                    try { return code.Lines[1, code.CountOfLines]; }
                    finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(code); }
                }
                Vbe.VBProject ReadProject(Excel.Workbook book) => book.VBProject;
                Office.DocumentProperties ReadProperties(Excel.Workbook book) =>
                    (Office.DocumentProperties)(object)book.CustomDocumentProperties;
                Office.DocumentInspector ReadInspector(Office.DocumentInspectors inspectors) => inspectors[1];
                Office.MsoDocInspectorStatus Inspect(Office.DocumentInspector inspector) {
                    inspector.Inspect(out Office.MsoDocInspectorStatus status, out string results);
                    return status;
                }
                Office.OfficeTheme ReadTheme(Excel.Workbook book) => book.Theme;
                Office.ThemeColor ReadColor(Office.ThemeColorScheme colors) =>
                    colors.Colors(Office.MsoThemeColorSchemeIndex.msoThemeAccent1);
                Office.ThemeFonts ReadMajorFonts(Office.ThemeFontScheme fonts) => fonts.MajorFont;
                Office.ThemeFont ReadFont(Office.ThemeFonts fonts) =>
                    fonts.Item(Office.MsoFontLanguageIndex.msoThemeLatin);
            }
            """).AddReferences(ExcelReference, OfficeReference, VbeReference);
        Assert.Empty(compilation.GetDiagnostics().Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        Assert.Empty(await compilation.WithAnalyzers([new SafetyAnalyzer()]).GetAnalyzerDiagnosticsAsync());

        using var output = new MemoryStream();
        var emitted = compilation.Emit(output);
        Assert.True(emitted.Success, string.Join(Environment.NewLine, emitted.Diagnostics));
        output.Position = 0;
        using var reader = new PEReader(output);
        var metadata = reader.GetMetadataReader();
        var references = metadata.AssemblyReferences.Select(handle =>
            metadata.GetString(metadata.GetAssemblyReference(handle).Name)).ToArray();
        Assert.DoesNotContain("office", references);
        Assert.DoesNotContain("Microsoft.Vbe.Interop", references);
        Assert.DoesNotContain("Microsoft.Office.Interop.Excel", references);
    }

    [Fact]
    public async Task TypedMetadata_ShapeFillReadDoesNotRequireOfficeDefinitions()
    {
        var compilation = CreateCompilation("""
            using Excel = Microsoft.Office.Interop.Excel;
            class C { object Read(Excel.Shape shape) => (object)shape.Fill; }
            """).AddReferences(ExcelReference);

        Assert.Empty(compilation.GetDiagnostics().Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        Assert.Empty(await compilation.WithAnalyzers([new SafetyAnalyzer()]).GetAnalyzerDiagnosticsAsync());
    }

    [Fact]
    public async Task TypedAdoDefinitions_CompileAndEmbedWithoutRuntimeAssemblyReferences()
    {
        var compilation = CreateCompilation("""
            using ADODB;
            using Excel = Microsoft.Office.Interop.Excel;
            class C {
                _Connection ReadConnection(Excel.ModelConnection connection) =>
                    (_Connection)connection.ADOConnection;
                _Recordset Execute(_Connection connection, string query) {
                    object affected;
                    return connection.Execute(query, out affected, (int)ExecuteOptionEnum.adOptionUnspecified);
                }
                _Recordset Schema(_Connection connection) =>
                    connection.OpenSchema(SchemaEnum.adSchemaCubes, System.Type.Missing, System.Type.Missing);
                object Read(_Recordset records) {
                    Fields fields = records.Fields;
                    Field field = fields[0];
                    try {
                        string name = field.Name;
                        object value = field.Value;
                        if (!records.EOF) records.MoveNext();
                        if (records.State == (int)ObjectStateEnum.adStateOpen) records.Close();
                        return value;
                    } finally {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(field);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(fields);
                    }
                }
                string Error(_Connection connection) {
                    Errors errors = connection.Errors;
                    ADODB.Error error = errors[0];
                    try { return $"{connection.ConnectionString}:{error.Number}:{error.Source}:{error.Description}"; }
                    finally {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(error);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(errors);
                    }
                }
            }
            """).AddReferences(ExcelReference, AdoReference);
        Assert.Empty(compilation.GetDiagnostics().Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        Assert.Empty(await compilation.WithAnalyzers([new SafetyAnalyzer()]).GetAnalyzerDiagnosticsAsync());

        var adoAssembly = Assert.IsAssignableFrom<IAssemblySymbol>(compilation.GetAssemblyOrModuleSymbol(AdoReference));
        Assert.Equal(new Version(2, 8, 0, 0), adoAssembly.Identity.Version);
        Assert.Contains(adoAssembly.GetAttributes(), attribute =>
            attribute.AttributeClass?.Name == "TypeLibVersionAttribute" &&
            attribute.ConstructorArguments is [{ Value: 2 }, { Value: 8 }]);
        Assert.Contains(adoAssembly.GetAttributes(), attribute =>
            attribute.AttributeClass?.Name == "GuidAttribute" &&
            string.Equals(attribute.ConstructorArguments[0].Value as string,
                "2A75196C-D9EB-4129-B803-931327F72D5C", StringComparison.OrdinalIgnoreCase));

        using var output = new MemoryStream();
        var emitted = compilation.Emit(output);
        Assert.True(emitted.Success, string.Join(Environment.NewLine, emitted.Diagnostics));
        output.Position = 0;
        using var reader = new PEReader(output);
        var metadata = reader.GetMetadataReader();
        var references = metadata.AssemblyReferences.Select(handle =>
            metadata.GetString(metadata.GetAssemblyReference(handle).Name)).ToArray();
        Assert.DoesNotContain("ADODB", references);
        Assert.DoesNotContain("Microsoft.Office.Interop.Excel", references);
    }

    [Theory]
    [InlineData("dynamic field;")]
    [InlineData("dynamic Property { get; set; }")]
    [InlineData("dynamic M(dynamic value) => value;")]
    [InlineData("void M() { dynamic value = null; }")]
    [InlineData("void M(System.Collections.Generic.List<dynamic> values) {}")]
    public async Task DynamicAccess_RejectsExplicitDeclarationTypes(string member)
    {
        var diagnostics = await AnalyzeAsync("class C { " + member + " }");

        Assert.Contains(diagnostics, diagnostic => diagnostic.Id == "EXL011");
    }

    [Fact]
    public async Task DynamicAccess_RejectsInferredLocalsAndDynamicCastsEvenWithComments()
    {
        const string source = """
            class C {
                void M(object value) {
                    // PIA gap: this does not establish a supported late-binding exception.
                    var lateBound = (dynamic)value;
                    lateBound.Refresh();
                }
            }
            """;
        var diagnostics = await AnalyzeAsync(source);

        Assert.Contains(diagnostics, diagnostic => diagnostic.Id == "EXL005");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Id == "EXL011");
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Id == "EXL011" &&
            source.Substring(diagnostic.Location.SourceSpan.Start, diagnostic.Location.SourceSpan.Length)
                .Contains("lateBound", StringComparison.Ordinal));
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

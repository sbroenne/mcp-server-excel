using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Testing;
using Microsoft.CodeAnalysis.Testing;
using Sbroenne.ExcelMcp.Analyzers;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("Feature", "PreCommit")]
[Trait("Feature", "AutomationSafety")]
[Trait("RequiresExcel", "false")]
public sealed class TypedSourceGuardTests
{
    private const string ResultTypes = """
        namespace Sbroenne.ExcelMcp.Core.Models
        {
            public class ResultBase
            {
                public bool Success { get; set; }
                public string ErrorMessage { get; set; }
            }
        }
        """;

    [Theory]
    [InlineData("{|#0:dynamic|} field = null;")]
    [InlineData("void Handle({|#0:dynamic|} pivot) {}")]
    [InlineData("{|#0:dynamic|} Find() => null;")]
    [InlineData("object Find(object source) => ({|#0:dynamic|})(object)source;")]
    [InlineData("System.Collections.Generic.List<{|#0:dynamic|}> items;")]
    [InlineData("{|#0:dynamic|}? field = null;")]
    public async Task DynamicGuard_RejectsPreviouslyMissedExecutableUses(string member)
    {
        var test = Create("class C { " + member + " }");
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "Use typed Excel PIA access"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("// TODO: finish migration")]
    [InlineData("// Reason: unavailable")]
    [InlineData("// PIA gap:")]
    [InlineData("// PIA gap: unavailable")]
    public async Task CompletedAreas_RejectGenericOrEmptyGapExemptions(string comment)
    {
        var test = Create("class C {\n" + comment + "\n{|#0:dynamic|} value;\n}");
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "verified runtime gap"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_AllowsVerifiedOccurrenceButNotNextOccurrence()
    {
        var test = Create("""
            class C
            {
                // PIA gap: Excel.Application.Run is unavailable to embedded PIAs.
                dynamic application;
                {|#0:dynamic|} unrelated;
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "adjacent"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_OfficeGapCannotExemptSupportedExcelMember()
    {
        var test = Create("""
            class C
            {
                object Read(Microsoft.Office.Interop.Excel.Range range)
                {
                    // PIA gap: Office.Core IDispatch getter.
                    return (({|#0:dynamic|})range).Value2;
                }
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "verified runtime gap"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_OfficeGapCannotExemptTypedRangeLocal()
    {
        var test = Create("""
            class C
            {
                void Read(Microsoft.Office.Interop.Excel.Range range)
                {
                    // PIA gap: Office.Core IDispatch getter.
                    {|#0:dynamic|} value = range;
                }
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "verified runtime gap"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("Font", "ThemeFont", "// PIA gap: Font.ThemeFont can return native null for differential formatting.")]
    [InlineData("Style", "IncludeFont", "// PIA gap: Style value-type getters can return native null.")]
    [InlineData("Font", "ThemeFont", "// PIA gap: Office.Core IDispatch getter.")]
    [InlineData("Style", "IncludeFont", "// PIA gap: Office.Core IDispatch getter.")]
    public async Task DynamicGuard_RejectsUnverifiedFormattingGaps(string type, string member, string comment)
    {
        var test = Create($$"""
            class C
            {
                object Read(Microsoft.Office.Interop.Excel.{{type}} value)
                {
                    {{comment}}
                    return (({|#0:dynamic|})value).{{member}};
                }
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "verified runtime gap"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_AllowsTypedFormattingAndVerifiedSourceNameSentinelGap()
    {
        var test = Create("""
            class C
            {
                object Theme(Microsoft.Office.Interop.Excel.Font font) => font.ThemeFont;
                bool Include(Microsoft.Office.Interop.Excel.Style style) => style.IncludeFont;
                object Source(Microsoft.Office.Interop.Excel.PivotField field)
                {
                    // PIA gap: typed SourceName is a string, but native #N/A is an integer sentinel.
                    return ((dynamic)field).SourceName;
                }
            }
            """);
        AddExcel(test);
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_RejectsRetiredNativeStyleIndentLevelWorkaround()
    {
        var test = Create("""
            class C
            {
                private static object? ReadNativeStyleIndentLevel(Microsoft.Office.Interop.Excel.Style style)
                {
                    // PIA gap: Style.IndentLevel returns VT_NULL (DBNull) for Normal and cleared-cell styles; its int getter throws InvalidCastException (0x80004002).
                    return (({|#0:dynamic|})style).IndentLevel;
                }
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "Use typed Excel PIA access"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("Microsoft.Office.Interop.Excel.Style", "IndentLevel", "// TODO: Style.IndentLevel returns VT_NULL.")]
    [InlineData("Microsoft.Office.Interop.Excel.Style", "IndentLevel", "// PIA gap: Style value-type getters can return native null.")]
    [InlineData("Microsoft.Office.Interop.Excel.Style", "IndentLevel", "// PIA gap: Office.Core IDispatch getter.")]
    [InlineData("Microsoft.Office.Interop.Excel.Style", "IncludeFont", "// PIA gap: Style.IndentLevel returns VT_NULL (DBNull).")]
    [InlineData("Microsoft.Office.Interop.Excel.Font", "ThemeFont", "// PIA gap: Style.IndentLevel returns VT_NULL (DBNull).")]
    [InlineData("Microsoft.Office.Interop.Excel.Range", "IndentLevel", "// PIA gap: Style.IndentLevel returns VT_NULL (DBNull).")]
    [InlineData("object", "IndentLevel", "// PIA gap: Style.IndentLevel returns VT_NULL (DBNull).")]
    public async Task DynamicGuard_RejectsStyleIndentLevelGapNearMisses(string receiver, string member, string comment)
    {
        var test = Create($$"""
            class C
            {
                object Read({{receiver}} value)
                {
                    {{comment}}
                    return (({|#0:dynamic|})value).{{member}};
                }
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "verified runtime gap"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_StyleIndentLevelCommentCannotExemptBlanketDynamicStyle()
    {
        var test = Create("""
            class C
            {
                object Read(Microsoft.Office.Interop.Excel.Style style)
                {
                    // PIA gap: Style.IndentLevel returns VT_NULL (DBNull).
                    {|#0:dynamic|} lateBound = style;
                    return lateBound.IndentLevel;
                }
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "adjacent"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_RetiredStyleIndentLevelCommentDoesNotExemptLaterUse()
    {
        var test = Create("""
            class C
            {
                object Read(Microsoft.Office.Interop.Excel.Style style)
                {
                    // PIA gap: Style.IndentLevel returns VT_NULL (DBNull).
                    int unrelated = 0;
                    return (({|#0:dynamic|})style).IndentLevel;
                }
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "adjacent"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_CommentsStringsAndIdentifiersAreNotTypes()
    {
        await Create("""
            class C
            {
                // dynamic? bad; ((dynamic)value).Rows.Count;
                string text = "dynamic ((dynamic)source)";
                int dynamic = 1;
            }
            """).RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_DynamicVariableNamedDynamicIsNotAnotherTypeOccurrence()
    {
        await Create("""
            class C
            {
                // PIA gap: Office.Core IDispatch object.
                dynamic @dynamic;
                object Read() => @dynamic.Value;
            }
            """).RunAsync();
    }

    [Fact]
    public async Task DynamicGuard_AssociatesGapWithElseBoundaryAndNestedCall()
    {
        await Create("""
            class C
            {
                void Write(object value)
                {
                    if (value == null) return;
                    // PIA gap: TextConnection.BackgroundQuery is absent from Excel PIA.
                    else ((dynamic)value).BackgroundQuery = false;
                }
                bool Read(object value)
                {
                    // PIA gap: TextConnection.BackgroundQuery is absent from Excel PIA.
                    return System.Convert.ToBoolean(((dynamic)value).BackgroundQuery);
                }
            }
            """).RunAsync();
    }

    [Fact]
    public async Task UnfinishedAreas_PreserveDocumentedDoubleCastOnly()
    {
        var test = Create("""
            class C
            {
                dynamic value;
                object Read(object source)
                {
                    // TODO: staged migration
                    return ((dynamic)source).Value;
                }
                object Bad(object source) => (({|#0:dynamic|})source).Value;
            }
            """, completed: false);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.DynamicId, 0, "PIA"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("int")]
    [InlineData("double")]
    [InlineData("System.DayOfWeek")]
    [InlineData("int?")]
    public async Task DynamicNumericCasts_RequireConvertEvenForVerifiedGap(string type)
    {
        var test = Create($$"""
            class C
            {
                // PIA gap: Office-core IDispatch property returns a numeric VARIANT.
                {{type}} Read(dynamic value) => {|#0:({{type}})value|};
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.NumericId, 0, "Convert.*"));
        await test.RunAsync();
    }

    [Fact]
    public async Task DynamicNumericConversions_AllowConvertAndTypedCasts()
    {
        await Create("""
            class C
            {
                // PIA gap: Office-core IDispatch numeric VARIANT.
                int Read(dynamic value) => System.Convert.ToInt32(value);
                int Typed(double value) => (int)value;
            }
            """).RunAsync();
    }

    [Theory]
    [InlineData("range.Rows.Count", "{|#0:range.Rows|}.Count")]
    [InlineData("table.Range.Row", "{|#0:table.Range|}.Row")]
    [InlineData("table.Range.Column", "{|#0:table.Range|}.Column")]
    public async Task ComGuard_UsesActualPiaTypesForChainedAcquisitions(string _, string marked)
    {
        var test = Create($$"""
            class C
            {
                int Read(Microsoft.Office.Interop.Excel.Range range, Microsoft.Office.Interop.Excel.ListObject table)
                    => {{marked}};
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.ComId, 0, "finally"));
        await test.RunAsync();
    }

    [Fact]
    public async Task ComGuard_RecognizesObjectReturningPiaPropertyCast()
    {
        var test = Create("""
            class C
            {
                string Read(Microsoft.Office.Interop.Excel.Range range)
                    => ({|#0:(Microsoft.Office.Interop.Excel.Worksheet)range.Parent|}).Name;
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.ComId, 0, "finally"));
        await test.RunAsync();
    }

    [Fact]
    public async Task ComGuard_RecognizesObjectReturningPiaPropertyThroughNestedCasts()
    {
        var test = Create("""
            using Excel = Microsoft.Office.Interop.Excel;
            class C
            {
                string Read(Excel.Range range)
                    => ({|#0:(Excel.Worksheet)(object)range.Parent|}).Name;
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.ComId, 0, "finally"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("System.Collections.Generic.List<Microsoft.Office.Interop.Excel.Range>", "ranges[0].Address")]
    [InlineData("System.Collections.Generic.IList<Microsoft.Office.Interop.Excel.Range>", "ranges[0].Address")]
    [InlineData("System.Collections.Generic.IReadOnlyList<Microsoft.Office.Interop.Excel.Range>", "ranges[0].Address")]
    [InlineData("System.Collections.Generic.Dictionary<int, Microsoft.Office.Interop.Excel.Range>", "ranges[0].Address")]
    [InlineData("Microsoft.Office.Interop.Excel.Range[]", "ranges[0].Address")]
    [InlineData("Holder", "ranges.Range.Address")]
    [InlineData("Holder", "ranges.GetRange().Address")]
    [InlineData("Holder", "((Microsoft.Office.Interop.Excel.Range)ranges.Value).Address")]
    [InlineData("System.Collections.Generic.List<Microsoft.Office.Interop.Excel.Range>", "((Microsoft.Office.Interop.Excel.Range)(object)ranges[0]).Address")]
    [InlineData("System.Collections.Generic.IList<Microsoft.Office.Interop.Excel.Range>", "((Microsoft.Office.Interop.Excel.Range)(object)ranges[0]).Address")]
    [InlineData("System.Collections.Generic.IReadOnlyList<Microsoft.Office.Interop.Excel.Range>", "((Microsoft.Office.Interop.Excel.Range)(object)ranges[0]).Address")]
    [InlineData("System.Collections.Generic.Dictionary<int, Microsoft.Office.Interop.Excel.Range>", "((Microsoft.Office.Interop.Excel.Range)(object)ranges[0]).Address")]
    [InlineData("Microsoft.Office.Interop.Excel.Range[]", "((Microsoft.Office.Interop.Excel.Range)(object)ranges[0]).Address")]
    [InlineData("Holder", "((Microsoft.Office.Interop.Excel.Range)(object)ranges.Range).Address")]
    [InlineData("Holder", "((Microsoft.Office.Interop.Excel.Range)(object)ranges.GetRange()).Address")]
    [InlineData("Holder", "((Microsoft.Office.Interop.Excel.Range)(object)ranges.Value).Address")]
    [InlineData("Holder", "((Microsoft.Office.Interop.Excel.Range)(object)ranges.GetValue()).Address")]
    public async Task ComGuard_AllowsAlreadyOwnedReferencesFromManagedHolders(string type, string expression)
    {
        var test = Create($$"""
            class Holder
            {
                public Microsoft.Office.Interop.Excel.Range Range { get; set; }
                public object Value { get; set; }
                public Microsoft.Office.Interop.Excel.Range GetRange() => Range;
                public object GetValue() => Value;
            }
            class C
            {
                string Read({{type}} ranges) => {{expression}};
            }
            """);
        AddExcel(test);
        await test.RunAsync();
    }

    [Theory]
    [InlineData("Microsoft.Office.Interop.Excel.Areas", "{|#0:ranges[1]|}.Address")]
    [InlineData("Microsoft.Office.Interop.Excel.Range", "({|#0:(Microsoft.Office.Interop.Excel.Range)ranges[1, 1]|}).Address")]
    [InlineData("System.Collections.Generic.List<Microsoft.Office.Interop.Excel.Range>", "{|#0:ranges[0].Rows|}.Address")]
    [InlineData("Microsoft.Office.Interop.Excel.Range", "({|#0:(Microsoft.Office.Interop.Excel.Range)(object)ranges[1, 1]|}).Address")]
    [InlineData("Microsoft.Office.Interop.Excel.Range", "((Microsoft.Office.Interop.Excel.Range)(object)({|#0:(Microsoft.Office.Interop.Excel.Range)ranges[1, 1]|})).Address")]
    [InlineData("Microsoft.Office.Interop.Excel.Areas", "((Microsoft.Office.Interop.Excel.Range)(object){|#0:ranges[1]|}).Address")]
    [InlineData("System.Collections.Generic.List<Microsoft.Office.Interop.Excel.Range>", "((Microsoft.Office.Interop.Excel.Range)(object){|#0:ranges[0].Rows|}).Address")]
    public async Task ComGuard_RejectsNativeAcquisitionsIncludingFromManagedHolders(string type, string expression)
    {
        var test = Create($$"""
            class C
            {
                string Read({{type}} ranges) => {{expression}};
            }
            """);
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.ComId, 0, "finally"));
        await test.RunAsync();
    }

    [Fact]
    public async Task ComGuard_DoesNotTraceNativeProvenanceThroughUserDefinedConversions()
    {
        await Create("""
            namespace Microsoft.Office.Interop.Excel
            {
                public class Range { public object Parent { get; } }
                public class Worksheet { public string Name { get; } }
            }
            class Wrapper
            {
                public static explicit operator Microsoft.Office.Interop.Excel.Worksheet(Wrapper value)
                    => new Microsoft.Office.Interop.Excel.Worksheet();
            }
            class C
            {
                string Read(Microsoft.Office.Interop.Excel.Range range)
                    => ((Microsoft.Office.Interop.Excel.Worksheet)(Wrapper)range.Parent).Name;
                string Nested(Microsoft.Office.Interop.Excel.Range range)
                    => ((Microsoft.Office.Interop.Excel.Worksheet)(object)(Microsoft.Office.Interop.Excel.Worksheet)(Wrapper)range.Parent).Name;
            }
            """).RunAsync();
    }

    [Theory]
    [InlineData("{|#0:anyName.Add(\"Example\", \"=Sheet1!A1\")|};")]
    [InlineData("_ = {|#0:anyName.Add(\"Example\", \"=Sheet1!A1\")|};")]
    public async Task ComGuard_RejectsDiscardedReturnsRegardlessOfVariableName(string statement)
    {
        var test = Create("class C { void Add(Microsoft.Office.Interop.Excel.Names anyName) { " + statement + " } }");
        AddExcel(test);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.ComId, 0, "Capture"));
        await test.RunAsync();
    }

    [Fact]
    public async Task ComGuard_AllowsManagedCollectionsAndCapturedAndBorrowedReferences()
    {
        var test = Create("""
            namespace Sbroenne.ExcelMcp.ComInterop.Session
            {
                public class ExcelContext { public Microsoft.Office.Interop.Excel.Workbook Book { get; set; } }
                public interface IExcelBatch { Microsoft.Office.Interop.Excel.Workbook GetWorkbook(string path); }
            }
            class ManagedResult { public System.Collections.Generic.List<int> Rows { get; } = new System.Collections.Generic.List<int>(); }
            class C
            {
                int Managed(ManagedResult result) => result.Rows.Count;
                string Borrowed(Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch) => batch.GetWorkbook("example").Name;
                void Read(Sbroenne.ExcelMcp.ComInterop.Session.ExcelContext ctx)
                {
                    var sheets = ctx.Book.Worksheets;
                    try { var count = sheets.Count; }
                    finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets); }
                }
            }
            """);
        AddExcel(test);
        await test.RunAsync();
    }

    [Fact]
    public async Task SuccessGuard_RejectsProvenSameResultAtReturnBeyondThirtyLines()
    {
        var test = Create(ResultTypes + """

            class C
            {
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read()
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                    result.Success = true;
            """ + new string('\n', 35) + """
                    result.ErrorMessage = "failure";
                    {|#0:return result;|}
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.SuccessId, 0, "nonempty"));
        await test.RunAsync();
    }

    [Fact]
    public async Task SuccessGuard_RejectsObjectInitializer()
    {
        var test = Create(ResultTypes + """

            class C
            {
                object Read() => {|#0:new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = "failure" }|};
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.SuccessId, 0, "clear the error"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("if (condition) result.ErrorMessage = \"first\"; else result.ErrorMessage = \"second\";", "return result;")]
    [InlineData("var alias = result; if (condition) alias.ErrorMessage = \"first\"; else result.ErrorMessage = \"second\";", "return alias;")]
    [InlineData("if (condition) { result.ErrorMessage = \"failure\"; {|#0:return result;|} }", "return new Sbroenne.ExcelMcp.Core.Models.ResultBase();")]
    public async Task SuccessGuard_RejectsProvenSimpleBranchConflicts(string mutations, string returned)
    {
        if (!mutations.Contains("{|#0:", StringComparison.Ordinal)) { returned = "{|#0:" + returned + "|}"; }
        var test = Create(ResultTypes + $$"""

            class C
            {
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read(bool condition)
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true };
                    {{mutations}}
                    {{returned}}
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.SuccessId, 0, "nonempty"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("if (condition) result.ErrorMessage = \"failure\"; else other.ErrorMessage = \"failure\";")]
    [InlineData("if (condition) { result.ErrorMessage = \"failure\"; result.Success = false; } else result.ErrorMessage = null;")]
    [InlineData("if (condition) Mutate(result); else result.ErrorMessage = \"failure\";")]
    [InlineData("while (condition) { result.ErrorMessage = \"failure\"; break; }")]
    [InlineData("var alias = result; if (condition) alias = other; alias.ErrorMessage = \"failure\";")]
    [InlineData("result.ErrorMessage = \"failure\"; var alias = result; if (condition) alias = other; alias.ErrorMessage = null;")]
    [InlineData("result.ErrorMessage = \"failure\"; var alias = result; if (condition) result = other; alias.ErrorMessage = null;")]
    public async Task SuccessGuard_AllowsDistinctRepairedOrUnknownBranchPaths(string mutations)
    {
        await Create(ResultTypes + $$"""

            class C
            {
                void Mutate(object value) {}
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read(bool condition)
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true };
                    var other = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                    {{mutations}}
                    return result;
                }
            }
            """).RunAsync();
    }

    [Theory]
    [InlineData("result.ErrorMessage = null;")]
    [InlineData("result.ErrorMessage = string.Empty;")]
    [InlineData("result.Success = false;")]
    [InlineData("var alias = result; alias.ErrorMessage = \"\";")]
    [InlineData("if (condition) result.ErrorMessage = null; else result.Success = false;")]
    [InlineData("result = new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = \"second\" }; result.ErrorMessage = null;")]
    [InlineData("Mutate(result);")]
    public async Task SuccessGuard_AllowsRepairedLocalInitializer(string repair)
    {
        await Create(ResultTypes + $$"""

            class C
            {
                void Mutate(object value) {}
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read(bool condition)
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = "failure" };
                    {{repair}}
                    return result;
                }
            }
            """).RunAsync();
    }

    [Fact]
    public async Task SuccessGuard_RejectsUnrepairedLocalInitializerAtReturn()
    {
        var test = Create(ResultTypes + """

            class C
            {
                object Read()
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = "failure" };
                    {|#0:return result;|}
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.SuccessId, 0, "nonempty"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SuccessGuard_TracksReferenceCastLocalInitializerUntilReturn(bool repaired)
    {
        var returned = repaired ? "return result;" : "{|#0:return result;|}";
        var test = Create(ResultTypes + $$"""

            class C
            {
                object Read()
                {
                    object result = (object)new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = "failure" };
                    {{(repaired ? "((Sbroenne.ExcelMcp.Core.Models.ResultBase)result).ErrorMessage = null;" : "")}}
                    {{returned}}
                }
            }
            """);
        if (!repaired) { test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.SuccessId, 0, "nonempty")); }
        await test.RunAsync();
    }

    [Theory]
    [InlineData("object result = {|#0:new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = \"failure\" }|};")]
    [InlineData("void Expose() { Accept({|#0:new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = \"failure\" }|}); }")]
    [InlineData("object Read() { return {|#0:new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = \"failure\" }|}; }")]
    [InlineData("object Read() { Sbroenne.ExcelMcp.Core.Models.ResultBase result; return result = {|#0:new Sbroenne.ExcelMcp.Core.Models.ResultBase { Success = true, ErrorMessage = \"failure\" }|}; }")]
    public async Task SuccessGuard_RejectsDirectlyExposedConflictingInitializer(string member)
    {
        var test = Create(ResultTypes + "class C { void Accept(object value) {} " + member + " }");
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.SuccessId, 0, "nonempty"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("var alias = result; alias.ErrorMessage = \"failure\"; result.Success = true;", "result")]
    [InlineData("var alias = result; result.ErrorMessage = \"failure\"; alias.Success = true;", "alias")]
    [InlineData("var alias = result; var next = alias; next.ErrorMessage = \"failure\"; alias.Success = true;", "next")]
    [InlineData("Sbroenne.ExcelMcp.Core.Models.ResultBase alias; alias = result; alias.ErrorMessage = \"failure\"; result.Success = true;", "alias")]
    [InlineData("var alias = result; result = new Sbroenne.ExcelMcp.Core.Models.ResultBase(); alias.ErrorMessage = \"failure\"; alias.Success = true;", "alias")]
    [InlineData("object alias = result; ((Sbroenne.ExcelMcp.Core.Models.ResultBase)alias).ErrorMessage = \"failure\"; result.Success = true;", "result")]
    [InlineData("object alias = (result); result.ErrorMessage = \"failure\"; ((Sbroenne.ExcelMcp.Core.Models.ResultBase)alias).Success = true;", "(Sbroenne.ExcelMcp.Core.Models.ResultBase)alias")]
    public async Task SuccessGuard_RejectsProvenConflictThroughLocalAliases(string mutations, string returned)
    {
        var test = Create(ResultTypes + $$"""

            class C
            {
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read()
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                    {{mutations}}
                    {|#0:return {{returned}};|}
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.SuccessId, 0, "nonempty"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData("alias.ErrorMessage = \"\";")]
    [InlineData("alias.ErrorMessage = string.Empty;")]
    [InlineData("alias.ErrorMessage = null;")]
    [InlineData("alias.Success = false;")]
    [InlineData("alias.Success ^= true;")]
    [InlineData("Mutate(alias);")]
    [InlineData("if (condition) alias.ErrorMessage = null;")]
    [InlineData("result = new Sbroenne.ExcelMcp.Core.Models.ResultBase(); result.Success = true;")]
    public async Task SuccessGuard_AllowsClearedOrUncertainAliasState(string change)
    {
        await Create(ResultTypes + $$"""

            class C
            {
                void Mutate(object value) {}
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read(bool condition)
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                    var alias = result;
                    alias.ErrorMessage = "failure";
                    result.Success = true;
                    {{change}}
                    return result;
                }
            }
            """).RunAsync();
    }

    [Theory]
    [InlineData("var alias = other; alias.ErrorMessage = \"failure\"; result.Success = true;")]
    [InlineData("var alias = result; alias = other; alias.ErrorMessage = \"failure\"; result.Success = true;")]
    [InlineData("var alias = result; alias = new Sbroenne.ExcelMcp.Core.Models.ResultBase(); alias.ErrorMessage = \"failure\"; result.Success = true;")]
    public async Task SuccessGuard_DoesNotMergeIndependentObjectsOrReassignedAliases(string mutations)
    {
        await Create(ResultTypes + $$"""

            class C
            {
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read()
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                    var other = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                    {{mutations}}
                    return result;
                }
            }
            """).RunAsync();
    }

    [Theory]
    [InlineData("result.ErrorMessage = \"\";")]
    [InlineData("result.ErrorMessage = string.Empty;")]
    [InlineData("result.ErrorMessage = null;")]
    [InlineData("result.Success = false;")]
    [InlineData("var other = new Sbroenne.ExcelMcp.Core.Models.ResultBase(); other.ErrorMessage = \"failure\"; result.ErrorMessage = null;")]
    [InlineData("if (condition) result.ErrorMessage = null;")]
    [InlineData("Mutate(result);")]
    [InlineData("var alias = result; alias.Success = false;")]
    public async Task SuccessGuard_DoesNotInferConflictsAcrossObjectsBranchesOrUnknownCalls(string change)
    {
        await Create(ResultTypes + $$"""

            class C
            {
                void Mutate(object value) {}
                Sbroenne.ExcelMcp.Core.Models.ResultBase Read(bool condition)
                {
                    var result = new Sbroenne.ExcelMcp.Core.Models.ResultBase();
                    result.Success = true;
                    result.ErrorMessage = "failure";
                    {{change}}
                    return result;
                }
            }
            """).RunAsync();
    }

    [Theory]
    [InlineData("\"xl/workbook.xml\"")]
    [InlineData("\"xl/worksheets/sheet1.xml\"")]
    [InlineData("\"[Content_Types].xml\"")]
    [InlineData("@\"xl\\connections.xml\"")]
    public async Task PackageGuard_RejectsWorkbookPartLiterals(string literal)
    {
        var test = Create("class C { string part = {|#0:" + literal + "|}; }");
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.PackageId, 0, "Excel COM"));
        await test.RunAsync();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PackageGuard_RejectsZipFedXmlIncludingHelperReturn(bool helper)
    {
        var expression = helper ? "Entry(archive)" : "archive.GetEntry(\"xl/\" + \"workbook.xml\").Open()";
        var test = Create($$"""
            class C
            {
                System.IO.Stream Entry(System.IO.Compression.ZipArchive archive) => archive.GetEntry("xl/" + "workbook.xml").Open();
                void Read(System.IO.Compression.ZipArchive archive)
                {
                    var document = {|#0:System.Xml.Linq.XDocument.Load({{expression}})|};
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.PackageId, 0, "ZIP"));
        await test.RunAsync();
    }

    [Fact]
    public async Task PackageGuard_RejectsZipFedXmlAcrossHelperParameter()
    {
        var test = Create("""
            class C
            {
                void Entry(System.IO.Compression.ZipArchive archive) => Read(archive.GetEntry("xl/" + "workbook.xml").Open());
                void Read(System.IO.Stream stream)
                {
                    var document = {|#0:System.Xml.Linq.XDocument.Load(stream)|};
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.PackageId, 0, "ZIP"));
        await test.RunAsync();
    }

    [Fact]
    public async Task PackageGuard_RejectsZipXmlReadThroughManagedTextReader()
    {
        var test = Create("""
            class C
            {
                void Read(System.IO.Compression.ZipArchive archive)
                {
                    using var reader = new System.IO.StreamReader(archive.GetEntry("xl/" + "workbook.xml").Open());
                    var xml = reader.ReadToEnd();
                    var document = {|#0:System.Xml.Linq.XDocument.Parse(xml)|};
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.PackageId, 0, "ZIP"));
        await test.RunAsync();
    }

    [Fact]
    public async Task PackageGuard_RejectsProjectAndImportedPartValues()
    {
        var test = Create("class C {}");
        test.TestState.AdditionalFiles.Add(("Runtime.csproj", "<Project>\n<PropertyGroup><Part>xl/workbook.xml</Part></PropertyGroup>\n</Project>"));
        test.ExpectedDiagnostics.Add(new DiagnosticResult(Assert.Single(new ExcelSafetyAnalyzer().SupportedDiagnostics,
            d => d.Id == ExcelSafetyAnalyzer.PackageId)).WithSpan("Runtime.csproj", 2, 17, 2, 17));
        await test.RunAsync();
    }

    [Fact]
    public async Task PackageGuard_AllowsSeparateZipXmlAndComments()
    {
        await Create("""
            class C
            {
                // xl/workbook.xml
                void Read(System.IO.Stream stream, string xml)
                {
                    using (var archive = new System.IO.Compression.ZipArchive(stream)) { var count = archive.Entries.Count; }
                    var document = System.Xml.Linq.XDocument.Parse(xml);
                }
            }
            """).RunAsync();
    }

    [Theory]
    [InlineData("archive.GetEntry(\"settings.xml\").Open()")]
    [InlineData("new System.IO.Compression.GZipStream(stream, System.IO.Compression.CompressionMode.Decompress)")]
    public async Task PackageGuard_AllowsUnrelatedCompressedXml(string source)
    {
        await Create($$"""
            class C
            {
                void Read(System.IO.Compression.ZipArchive archive, System.IO.Stream stream)
                {
                    using var input = {{source}};
                    var document = System.Xml.Linq.XDocument.Load(input);
                }
            }
            """).RunAsync();
    }

    [Fact]
    public async Task PackageGuard_RejectsKnownWorkbookArchiveThroughHelpers()
    {
        var test = Create("""
            class C
            {
                System.IO.Stream Entry(System.IO.Compression.ZipArchive archive) => archive.Entries[0].Open();
                void Read(System.IO.Stream stream)
                {
                    var document = {|#0:System.Xml.Linq.XDocument.Load(stream)|};
                }
                void Workbook()
                {
                    using var archive = System.IO.Compression.ZipFile.OpenRead("report.xlsx");
                    Read(Entry(archive));
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.PackageId, 0, "ZIP"));
        await test.RunAsync();
    }

    [Fact]
    public async Task PackageGuard_RejectsKnownWorkbookStreamThroughArchiveConstructor()
    {
        var test = Create("""
            class C
            {
                System.IO.Stream Workbook(string path) => System.IO.File.OpenRead(path);
                void Read()
                {
                    using var stream = Workbook("report.xlsx");
                    using var archive = new System.IO.Compression.ZipArchive(stream);
                    var document = {|#0:System.Xml.Linq.XDocument.Load(archive.Entries[0].Open())|};
                }
            }
            """);
        test.ExpectedDiagnostics.Add(Expected(ExcelSafetyAnalyzer.PackageId, 0, "ZIP"));
        await test.RunAsync();
    }

    internal static CSharpAnalyzerTest<ExcelSafetyAnalyzer, DefaultVerifier> Create(string source, bool completed = true)
    {
        var test = new CSharpAnalyzerTest<ExcelSafetyAnalyzer, DefaultVerifier>
        {
            ReferenceAssemblies = ReferenceAssemblies.Net.Net80,
            CompilerDiagnostics = CompilerDiagnostics.Errors
        };
        test.TestState.Sources.Add(("/0/Cases.cs", source));
        test.TestState.AnalyzerConfigFiles.Add(("/0/.editorconfig", "root = true\n[*.cs]\nexcelmcp_typed_excel = " + completed.ToString().ToLowerInvariant()));
        test.SolutionTransforms.Add((solution, projectId) =>
            solution.WithProjectParseOptions(projectId, new CSharpParseOptions(LanguageVersion.Preview)));
        return test;
    }

    private static DiagnosticResult Expected(string id, int marker, string guidance)
    {
        var descriptor = Assert.Single(new ExcelSafetyAnalyzer().SupportedDiagnostics, d => d.Id == id);
        Assert.Contains(guidance, descriptor.MessageFormat.ToString(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal);
        return new DiagnosticResult(descriptor).WithLocation(marker);
    }

    private static void AddExcel(CSharpAnalyzerTest<ExcelSafetyAnalyzer, DefaultVerifier> test)
    {
        var root = FindRoot();
        var assets = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "tests", "ExcelMcp.ScriptSafety.Tests", "obj", "project.assets.json")));
        using (assets)
        {
            var folder = assets.RootElement.GetProperty("packageFolders").EnumerateObject().First().Name;
            var package = assets.RootElement.GetProperty("libraries").EnumerateObject()
                .Single(p => p.Name.StartsWith("Microsoft.Office.Interop.Excel/", StringComparison.OrdinalIgnoreCase)).Name;
            test.TestState.AdditionalReferences.Add(MetadataReference.CreateFromFile(Path.Combine(folder,
                package.Replace('/', Path.DirectorySeparatorChar).ToLowerInvariant(), "lib", "netstandard2.0", "Microsoft.Office.Interop.Excel.dll")));
        }
    }

    internal static string FindRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { directory = directory.Parent; }
        return directory?.FullName ?? throw new DirectoryNotFoundException("Repository root not found.");
    }
}

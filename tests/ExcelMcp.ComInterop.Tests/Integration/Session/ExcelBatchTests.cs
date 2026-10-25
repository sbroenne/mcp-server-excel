using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelBatch - verifies batch operations and COM cleanup.
/// Tests that Excel instances are reused across operations and properly cleaned up.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test ExcelBatch.Execute() reuses Excel instance
/// - ✅ Test ExcelBatch.Dispose() COM cleanup
/// - ✅ Test ExcelBatch.Save() functionality
/// - ✅ Verify Excel.exe process termination (no leaks)
///
/// NOTE: ExcelBatch.Dispose() handles all GC cleanup automatically.
/// Tests only need to wait for async disposal and process termination timing.
///
/// These tests run sequentially and verify only their owned Excel processes.
/// Startup-hook and configured native-capability controls run OnDemand.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
[Trait("RequiresExcel", "true")]
public partial class ExcelBatchTests : IAsyncLifetime, IDisposable
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;
    private readonly OwnedExcelProcessScope _owned = new();
    private readonly List<string> _temporaryFiles = new();
    private readonly List<ExcelProcessIdentity> _startupIdentities = new();

    private static string? GetConfiguredIrmTestFilePath()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        return !string.IsNullOrWhiteSpace(irmTestFile) && File.Exists(irmTestFile)
            ? Path.GetFullPath(irmTestFile)
            : null;
    }

    public ExcelBatchTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        // Use static test file from TestFiles folder (must be pre-created)
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.xlsx");

            // Verify the static file exists
            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}. " +
                    "Please create the batch-test-static.xlsx file in the TestFiles folder.");
            }
        }

        // Create a fresh copy for this test instance (in temp folder)
        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-test-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(_testFileCopy);
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        return Task.CompletedTask;
    }

    public Task DisposeAsync()
    {
        Dispose();
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        var failures = new List<Exception>();
        foreach (var identity in _startupIdentities)
        {
            var failure = Record.Exception(() =>
            {
                Assert.True(SpinWait.SpinUntil(() => OwnedProcessGuard.TryConfirmExited(identity),
                    ProcessTerminationPolicy.NormalGraceTimeout),
                    $"Excel {identity} is still running after failed startup.");
                SessionManager.UntrackExcelProcess(identity);
            });
            if (failure is not null) { failures.Add(failure); }
        }
        var cleanup = Record.Exception(() =>
            SessionTestCleanup.AssertExitedAndDelete(_owned, _temporaryFiles));
        if (cleanup is not null) { failures.Add(cleanup); }
        if (failures.Count > 0)
        {
            throw new AggregateException("ExcelBatch test cleanup failed.", failures);
        }
    }

    [Fact]
    public void Execute_EmbeddedOfficeDefinitions_ReadSecurityAndDrawingState()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Assert.Equal(Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable, ctx.App.AutomationSecurity);
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.FillFormat? fill = null;
            Excel.ColorFormat? color = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                shapes = sheet.Shapes;
                int originalCount = shapes.Count;
                shape = shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 20, 30, 80, 40);
                shape.Name = "EmbeddedDefinitionsProbe";
                fill = shape.Fill;
                fill.Solid();
                color = fill.ForeColor;
                color.RGB = 0x563412;

                Assert.Equal(originalCount + 1, shapes.Count);
                Assert.Equal("EmbeddedDefinitionsProbe", shape.Name);
                Assert.Equal(Office.MsoShapeType.msoAutoShape, shape.Type);
                Assert.Equal(Office.MsoFillType.msoFillSolid, fill.Type);
                Assert.Equal(0x563412, color.RGB);
                Assert.Equal(80f, shape.Width);
                Assert.Equal(40f, shape.Height);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref color);
                ComUtilities.Release(ref fill);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Fact]
    public void Execute_EmbeddedOfficeDefinitions_ReadThemesAndInspectors()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Office.OfficeTheme? theme = null;
            Office.ThemeColorScheme? colors = null;
            Office.ThemeColor? accent = null;
            Office.ThemeFontScheme? fonts = null;
            Office.ThemeFonts? majorFonts = null;
            Office.ThemeFonts? minorFonts = null;
            Office.DocumentInspectors? inspectors = null;
            try
            {
                theme = ctx.Book.Theme;
                colors = theme.ThemeColorScheme;
                accent = colors.Colors(Office.MsoThemeColorSchemeIndex.msoThemeAccent1);
                accent.RGB = 0x563412;
                Assert.Equal(0x563412, accent.RGB);
                fonts = theme.ThemeFontScheme;
                majorFonts = fonts.MajorFont;
                minorFonts = fonts.MinorFont;
                string[] majorNames = ["Cambria", "Tahoma", "Yu Mincho"];
                string[] minorNames = ["Calibri", "Arial", "Yu Gothic"];
                for (int index = 1; index <= majorNames.Length; index++)
                {
                    Office.ThemeFont? major = null;
                    Office.ThemeFont? minor = null;
                    try
                    {
                        var language = (Office.MsoFontLanguageIndex)index;
                        major = majorFonts.Item(language);
                        minor = minorFonts.Item(language);
                        major.Name = majorNames[index - 1];
                        minor.Name = minorNames[index - 1];
                        Assert.Equal(majorNames[index - 1], major.Name);
                        Assert.Equal(minorNames[index - 1], minor.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref minor);
                        ComUtilities.Release(ref major);
                    }
                }

                inspectors = ctx.Book.DocumentInspectors;
                Assert.True(inspectors.Count > 0);
                for (int index = 1; index <= inspectors.Count; index++)
                {
                    Office.DocumentInspector? inspector = null;
                    try
                    {
                        inspector = inspectors[index];
                        Assert.False(string.IsNullOrWhiteSpace(inspector.Name));
                        inspector.Inspect(out Office.MsoDocInspectorStatus status, out string results);
                        Assert.True(status is Office.MsoDocInspectorStatus.msoDocInspectorStatusDocOk
                            or Office.MsoDocInspectorStatus.msoDocInspectorStatusIssueFound,
                            $"Inspector '{inspector.Name}' returned {status}: {results}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref inspector);
                    }
                }
                Assert.Equal(0x563412, accent.RGB);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref inspectors);
                ComUtilities.Release(ref minorFonts);
                ComUtilities.Release(ref majorFonts);
                ComUtilities.Release(ref fonts);
                ComUtilities.Release(ref accent);
                ComUtilities.Release(ref colors);
                ComUtilities.Release(ref theme);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_DocumentPropertyCollections_DoNotExposeOfficeDefinitionInterface(bool builtIn)
    {
        Assert.Equal(typeof(ExcelBatchTests).Assembly, typeof(Office.DocumentProperties).Assembly);
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            object? properties = null;
            try
            {
                properties = builtIn ? ctx.Book.BuiltinDocumentProperties : ctx.Book.CustomDocumentProperties;
                Assert.NotNull(properties);
                var error = Assert.Throws<InvalidCastException>(() => (Office.DocumentProperties)properties);
                Assert.Equal(unchecked((int)0x80004002), error.HResult);
                _output.WriteLine(
                    $"{(builtIn ? "Built-in" : "Custom")} document properties rejected Office.DocumentProperties " +
                    $"({typeof(Office.DocumentProperties).GUID}) with 0x{error.HResult:X8}.");
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref properties);
            }
        });
    }

    [Fact]
    public void Execute_EmbeddedVbeDefinitions_ReadAndWriteModuleSource()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Vbe.VBProject? project = null;
            Vbe.VBComponents? components = null;
            Vbe.VBComponent? component = null;
            Vbe.CodeModule? code = null;
            try
            {
                project = ctx.Book.VBProject;
                Assert.False(string.IsNullOrWhiteSpace(project.Name));
                Assert.Equal(Vbe.vbext_ProjectProtection.vbext_pp_none, project.Protection);
                Assert.Equal(Vbe.vbext_VBAMode.vbext_vm_Design, project.Mode);
                components = project.VBComponents;
                int originalCount = components.Count;
                component = components.Add(Vbe.vbext_ComponentType.vbext_ct_StdModule);
                component.Name = "EmbeddedDefinitionsProbe";
                code = component.CodeModule;
                code.AddFromString("Public Sub EmbeddedProbe()\r\nEnd Sub");

                Assert.Equal(originalCount + 1, components.Count);
                Assert.Equal("EmbeddedDefinitionsProbe", component.Name);
                Assert.Equal(Vbe.vbext_ComponentType.vbext_ct_StdModule, component.Type);
                Assert.Equal(2, code.CountOfLines);
                Assert.Equal("Public Sub EmbeddedProbe()", code.Lines[1, 1]);
                Assert.Equal("End Sub", code.Lines[2, 1]);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref code);
                ComUtilities.Release(ref component);
                ComUtilities.Release(ref components);
                ComUtilities.Release(ref project);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Fact]
    public void Execute_EmbeddedAdoDefinitions_ReadLiveModelWithoutRuntimeInteropDll()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            var originalSecurity = ctx.App.AutomationSecurity;
            Vbe.VBProject? project = null;
            Vbe.VBComponents? components = null;
            Vbe.VBComponent? component = null;
            Vbe.CodeModule? code = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? source = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? tableConnection = null;
            Excel.Model? model = null;
            Excel.ModelTables? modelTables = null;
            Excel.WorkbookConnection? modelConnection = null;
            Excel.ModelConnection? modelSession = null;
            ADODB._Connection? ado = null;
            ADODB._Recordset? records = null;
            ADODB.Fields? fields = null;
            ADODB.Field? field = null;
            ADODB.Errors? errors = null;
            ADODB.Error? error = null;
            Exception? operationFailure = null;
            void CloseRecords()
            {
                var failures = new List<Exception>();
                if (operationFailure is not null) failures.Add(operationFailure);
                try
                {
                    if (records is not null && (records.State & (int)ADODB.ObjectStateEnum.adStateOpen) != 0)
                        records.Close();
                }
                catch (Exception cleanupFailure)
                {
                    failures.Add(cleanupFailure);
                }
                try
                {
                    ctx.App.AutomationSecurity = originalSecurity;
                }
                catch (Exception cleanupFailure)
                {
                    failures.Add(cleanupFailure);
                }
                if (failures.Count > (operationFailure is null ? 0 : 1))
                {
                    throw new AggregateException("ADO definition check or cleanup failed.", failures);
                }
            }
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                source = sheet.Range["A1:A3"];
                source.Value2 = new object[,] { { "Value" }, { 7 }, { 11 } };
                tables = sheet.ListObjects;
                table = tables.Add(Excel.XlListObjectSourceType.xlSrcRange, source,
                    Type.Missing, Excel.XlYesNoGuess.xlYes);
                table.Name = "AdoEmbeddingTable";
                connections = ctx.Book.Connections;
                tableConnection = connections.Add2("AdoEmbeddingConnection", "ADO definition check",
                    $"WORKSHEET;{ctx.Book.Path}", $"{ctx.Book.Name}!{table.Name}",
                    Excel.XlCmdType.xlCmdExcel, true, false);
                model = ctx.Book.Model;
                modelTables = model.ModelTables;
                Assert.Equal(1, modelTables.Count);
                modelConnection = model.DataModelConnection;
                modelSession = modelConnection.ModelConnection;
                ado = ComUtilities.ReadAdoConnection(modelSession);
                Assert.NotNull(ado);
                Assert.Equal((int)ADODB.ObjectStateEnum.adStateOpen, ado.State);
                Assert.Contains("Provider=", ado.ConnectionString, StringComparison.OrdinalIgnoreCase);
                project = ctx.Book.VBProject;
                components = project.VBComponents;
                component = components.Add(Vbe.vbext_ComponentType.vbext_ct_StdModule);
                code = component.CodeModule;
                code.AddFromString("""
                    Public Function NativeAdoError(ByVal connection As Object) As Variant
                        Dim rows As Object
                        connection.Errors.Clear
                        On Error Resume Next
                        Set rows = connection.Execute("EVALUATE MissingTableForAdoDefinitions")
                        Dim result(1 To 1, 1 To 6) As Variant
                        result(1, 1) = Err.Number
                        result(1, 2) = Err.Description
                        On Error GoTo 0
                        result(1, 3) = connection.Errors.Count
                        If connection.Errors.Count > 0 Then
                            result(1, 4) = connection.Errors.Item(0).Number
                            result(1, 5) = connection.Errors.Item(0).Description
                            result(1, 6) = connection.Errors.Item(0).Source
                        End If
                        NativeAdoError = result
                    End Function
                    """);

                for (int attempt = 0; attempt < 2; attempt++)
                {
                    records = ado.Execute(
                        "EVALUATE ROW(\"Total\", SUM('AdoEmbeddingTable'[Value]), \"Label\", \"checked\")",
                        out _, (int)ADODB.ExecuteOptionEnum.adOptionUnspecified);
                    fields = records.Fields;
                    Assert.Equal(2, fields.Count);
                    Assert.False(records.EOF);
                    field = fields[0];
                    Assert.Equal("[Total]", field.Name);
                    object total = field.Value;
                    Assert.Equal(18, Convert.ToInt32(total, CultureInfo.InvariantCulture));
                    ComUtilities.Release(ref field);
                    field = fields[1];
                    Assert.Equal("[Label]", field.Name);
                    Assert.Equal("checked", (object)field.Value);
                    ComUtilities.Release(ref field);
                    records.MoveNext();
                    Assert.True(records.EOF);
                    records.Close();
                    Assert.Equal((int)ADODB.ObjectStateEnum.adStateClosed, records.State);
                    ComUtilities.Release(ref fields);
                    ComUtilities.Release(ref records);

                    if (attempt == 0)
                    {
                        errors = ado.Errors;
                        errors.Clear();
                        ComUtilities.Release(ref errors);
                        var queryFailure = Assert.Throws<COMException>(() =>
                        {
                            records = ado.Execute("EVALUATE MissingTableForAdoDefinitions",
                                out _, (int)ADODB.ExecuteOptionEnum.adOptionUnspecified);
                        });
                        errors = ado.Errors;
                        int errorCount = errors.Count;
                        int? errorNumber = null;
                        string? description = null;
                        string? errorSource = null;
                        if (errorCount > 0)
                        {
                            error = errors[0];
                            errorNumber = error.Number;
                            description = error.Description;
                            errorSource = error.Source;
                            Assert.NotEqual(0, errorNumber);
                        }
                        ctx.App.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityLow;
                        object? native = ctx.App.Run($"'{ctx.Book.Name}'!NativeAdoError", ado);
                        ctx.App.AutomationSecurity = originalSecurity;
                        var expected = Assert.IsType<object[,]>(native);
                        Assert.Equal(Convert.ToInt32(expected[1, 1], CultureInfo.InvariantCulture), queryFailure.HResult);
                        // VBA and .NET supply different wrapper text; preserve the native code and provider records.
                        Assert.False(string.IsNullOrWhiteSpace((string)expected[1, 2]));
                        Assert.False(string.IsNullOrWhiteSpace(queryFailure.Message));
                        Assert.Equal(Convert.ToInt32(expected[1, 3], CultureInfo.InvariantCulture), errorCount);
                        if (errorCount > 0)
                        {
                            Assert.Equal(Convert.ToInt32(expected[1, 4], CultureInfo.InvariantCulture), errorNumber);
                            Assert.Equal(expected[1, 5], description ?? string.Empty);
                            Assert.Equal(expected[1, 6], errorSource ?? string.Empty);
                        }
                        ComUtilities.Release(ref error);
                        ComUtilities.Release(ref errors);
                    }
                }

                records = ado.OpenSchema(ADODB.SchemaEnum.adSchemaCubes, Type.Missing, Type.Missing);
                fields = records.Fields;
                Assert.False(records.EOF);
                field = fields["CUBE_NAME"];
                Assert.Equal("Model", (object)field.Value);
                Assert.Equal((int)ADODB.ObjectStateEnum.adStateOpen, ado.State);
                return 0;
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                try
                {
                    CloseRecords();
                }
                finally
                {
                    ComUtilities.Release(ref error);
                    ComUtilities.Release(ref errors);
                    ComUtilities.Release(ref code);
                    ComUtilities.Release(ref component);
                    ComUtilities.Release(ref components);
                    ComUtilities.Release(ref project);
                    ComUtilities.Release(ref field);
                    ComUtilities.Release(ref fields);
                    ComUtilities.Release(ref records);
                    ComUtilities.Release(ref ado);
                    ComUtilities.Release(ref modelSession);
                    ComUtilities.Release(ref modelConnection);
                    ComUtilities.Release(ref modelTables);
                    ComUtilities.Release(ref model);
                    ComUtilities.Release(ref tableConnection);
                    ComUtilities.Release(ref connections);
                    ComUtilities.Release(ref table);
                    ComUtilities.Release(ref tables);
                    ComUtilities.Release(ref source);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Fact]
    public void Execute_TypedFormattingDefinitions_PreserveMixedRangeValues()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? first = null;
            Excel.Range? second = null;
            Excel.Range? combined = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                first = sheet.Range["A1"];
                second = sheet.Range["A2"];
                combined = sheet.Range["A1:A2"];
                first.WrapText = true;
                second.WrapText = false;
                first.Locked = true;
                second.Locked = false;
                first.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                second.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                first.NumberFormat = "0.00";
                second.NumberFormat = "0%";

                object? wrap = combined.WrapText;
                object? locked = combined.Locked;
                object? alignment = combined.HorizontalAlignment;
                object? numberFormat = combined.NumberFormat;
                Assert.IsType<DBNull>(wrap);
                Assert.IsType<DBNull>(locked);
                Assert.IsType<DBNull>(alignment);
                Assert.IsType<DBNull>(numberFormat);
                Assert.True((bool)(object)first.WrapText);
                Assert.False((bool)(object)second.WrapText);
                Assert.True((bool)(object)first.Locked);
                Assert.False((bool)(object)second.Locked);
                Assert.Equal((int)Excel.XlHAlign.xlHAlignLeft, (int)(object)first.HorizontalAlignment);
                Assert.Equal((int)Excel.XlHAlign.xlHAlignRight, (int)(object)second.HorizontalAlignment);
                Assert.Equal("0.00", (string)(object)first.NumberFormat);
                Assert.Equal("0%", (string)(object)second.NumberFormat);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref combined);
                ComUtilities.Release(ref second);
                ComUtilities.Release(ref first);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Theory]
    [InlineData(false, null)]
    [InlineData(false, 0)]
    [InlineData(false, 2)]
    [InlineData(true, null)]
    public void Execute_TypedFormattingDefinitions_IdentifyUnsetStyleLimits(bool differentialFont, int? specified)
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, _) =>
        {
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            Excel.TableStyles? tableStyles = null;
            Excel.TableStyle? tableStyle = null;
            Excel.TableStyleElements? elements = null;
            Excel.TableStyleElement? element = null;
            Excel.Font? font = null;
            Vbe.VBProject? project = null;
            Vbe.VBComponents? components = null;
            Vbe.VBComponent? module = null;
            Vbe.CodeModule? code = null;
            Office.MsoAutomationSecurity originalSecurity = ctx.App.AutomationSecurity;
            var failures = new List<Exception>();
            try
            {
                styles = ctx.Book.Styles;
                style = styles.Add("UnsetDefinitionProbe");
                if (specified.HasValue) { style.IndentLevel = specified.Value; }
                style.IncludeAlignment = false;
                style.IncludeProtection = false;
                tableStyles = ctx.Book.TableStyles;
                tableStyle = tableStyles.Add("UnsetDefinitionProbe");
                elements = tableStyle.TableStyleElements;
                element = elements.Item(Excel.XlTableStyleElementType.xlWholeTable);
                font = element.Font;
                Assert.False(element.HasFormat);
                Assert.False(style.IncludeAlignment);
                Assert.False(style.IncludeProtection);

                project = ctx.Book.VBProject;
                components = project.VBComponents;
                module = components.Add(Vbe.vbext_ComponentType.vbext_ct_StdModule);
                module.Name = "NativeFormattingProbe";
                code = module.CodeModule;
                // VBA supplies the native values for comparison, not a product fallback.
                code.AddFromString("""
                    Public Function ReadStyleVariants() As Variant
                        Dim values(1 To 2, 1 To 10) As Variant
                        Dim definition As Object
                        Set definition = ThisWorkbook.Styles("UnsetDefinitionProbe")
                        On Error Resume Next
                        values(1, 1) = definition.HorizontalAlignment
                        values(2, 1) = Err.Number: Err.Clear
                        values(1, 2) = definition.VerticalAlignment
                        values(2, 2) = Err.Number: Err.Clear
                        values(1, 3) = definition.WrapText
                        values(2, 3) = Err.Number: Err.Clear
                        values(1, 4) = definition.ShrinkToFit
                        values(2, 4) = Err.Number: Err.Clear
                        values(1, 5) = definition.AddIndent
                        values(2, 5) = Err.Number: Err.Clear
                        values(1, 6) = definition.IndentLevel
                        values(2, 6) = Err.Number: Err.Clear
                        values(1, 7) = definition.Orientation
                        values(2, 7) = Err.Number: Err.Clear
                        values(1, 8) = definition.ReadingOrder
                        values(2, 8) = Err.Number: Err.Clear
                        values(1, 9) = definition.Locked
                        values(2, 9) = Err.Number: Err.Clear
                        values(1, 10) = definition.FormulaHidden
                        values(2, 10) = Err.Number
                        ReadStyleVariants = values
                    End Function
                    Public Function ReadThemeFontVariant() As Variant
                        Dim values(1 To 2, 1 To 1) As Variant
                        Dim definition As TableStyle
                        Dim element As TableStyleElement
                        Dim definitionFont As Font
                        Set definition = ThisWorkbook.TableStyles("UnsetDefinitionProbe")
                        Set element = definition.TableStyleElements(xlWholeTable)
                        Set definitionFont = element.Font
                        On Error Resume Next
                        values(1, 1) = definitionFont.ThemeFont
                        values(2, 1) = Err.Number
                        ReadThemeFontVariant = values
                    End Function
                    """);
                ctx.App.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityLow;
                string procedure = differentialFont ? "ReadThemeFontVariant" : "ReadStyleVariants";
                object? native = ctx.App.Run($"'{ctx.Book.Name.Replace("'", "''", StringComparison.Ordinal)}'!NativeFormattingProbe.{procedure}");
                var expected = Assert.IsType<object[,]>(native);
                Func<object?>[] reads = differentialFont
                    ? [() => (int)font.ThemeFont]
                    :
                    [
                        () => (int)style.HorizontalAlignment, () => (int)style.VerticalAlignment,
                        () => style.WrapText, () => style.ShrinkToFit, () => style.AddIndent,
                        () => style.IndentLevel, () => (int)style.Orientation, () => style.ReadingOrder,
                        () => style.Locked, () => style.FormulaHidden
                    ];
                for (int index = 0; index < reads.Length; index++)
                {
                    int column = index + 1;
                    int nativeError = Convert.ToInt32(expected[2, column], CultureInfo.InvariantCulture);
                    _output.WriteLine($"Native field {column}: {expected[1, column] ?? "<null>"} ({expected[1, column]?.GetType().Name ?? "null"}), error {nativeError}");
                    object? actual = null;
                    var failure = Record.Exception(() => actual = reads[index]());
                    _output.WriteLine($"Typed field {column}: {actual ?? "<null>"}, failure {failure?.GetType().Name ?? "none"}, HRESULT 0x{failure?.HResult ?? 0:X8}");
                    if (!differentialFont && column == 6 && expected[1, column] is DBNull)
                    {
                        Assert.Equal(0, nativeError);
                        var conversion = Assert.IsType<InvalidCastException>(failure);
                        Assert.Equal(unchecked((int)0x80004002), conversion.HResult);
                        Assert.Null(ComUtilities.ReadStyleIndentLevel(style));
                    }
                    else if (nativeError == 0)
                    {
                        Assert.Null(failure);
                        Assert.Equal(expected[1, column], actual);
                    }
                    else if (nativeError == 94)
                    {
                        Assert.IsType<InvalidCastException>(failure);
                    }
                    else
                    {
                        Assert.Equal(438, nativeError);
                        var nativeFailure = Assert.IsType<COMException>(failure);
                        Assert.True(ComErrors.IsMemberUnavailable(nativeFailure),
                            $"Unexpected typed getter failure: 0x{nativeFailure.HResult:X8} {nativeFailure.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                failures.Add(ex);
            }
            finally
            {
                try
                {
                    var restore = Record.Exception(() => ctx.App.AutomationSecurity = originalSecurity);
                    if (restore is not null) { failures.Add(restore); }
                }
                finally
                {
                    ComUtilities.Release(ref code);
                    ComUtilities.Release(ref module);
                    ComUtilities.Release(ref components);
                    ComUtilities.Release(ref project);
                    ComUtilities.Release(ref font);
                    ComUtilities.Release(ref element);
                    ComUtilities.Release(ref elements);
                    ComUtilities.Release(ref tableStyle);
                    ComUtilities.Release(ref tableStyles);
                    ComUtilities.Release(ref style);
                    ComUtilities.Release(ref styles);
                }
            }
            if (failures.Count > 0)
                throw new AggregateException("Formatting comparison or security restoration failed.", failures);
            Assert.Equal(originalSecurity, ctx.App.AutomationSecurity);
            return 0;
        });
        AssertNoExternalInteropAssemblies();
    }

    [Theory]
    [InlineData(null, false, false)]
    [InlineData(null, false, true)]
    [InlineData(null, true, false)]
    [InlineData(null, true, true)]
    [InlineData(0, false, false)]
    [InlineData(0, false, true)]
    [InlineData(0, true, false)]
    [InlineData(0, true, true)]
    [InlineData(2, false, false)]
    [InlineData(2, false, true)]
    [InlineData(2, true, false)]
    [InlineData(2, true, true)]
    public void Execute_StyleIndentLevel_PreservesUnsetAndExplicitValues(
        int? indentLevel, bool includeAlignment, bool includeProtection)
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, _) =>
        {
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            try
            {
                styles = ctx.Book.Styles;
                style = styles.Add("IndentLevelProbe");
                if (indentLevel.HasValue) { style.IndentLevel = indentLevel.Value; }
                style.IncludeAlignment = includeAlignment;
                style.IncludeProtection = includeProtection;

                int? expected = indentLevel is null or 0 ? null : indentLevel;
                if (!expected.HasValue)
                {
                    var failure = Assert.Throws<InvalidCastException>(() => style.IndentLevel);
                    Assert.Equal(unchecked((int)0x80004002), failure.HResult);
                }
                Assert.Equal(expected, ComUtilities.ReadStyleIndentLevel(style));
                if (indentLevel == 0)
                {
                    style.IndentLevel = 2;
                    Assert.Equal(2, ComUtilities.ReadStyleIndentLevel(style));
                    style.IndentLevel = 0;
                    Assert.Equal(0, style.IndentLevel);
                    Assert.Equal(0, ComUtilities.ReadStyleIndentLevel(style));
                }
                Assert.Equal(includeAlignment, style.IncludeAlignment);
                Assert.Equal(includeProtection, style.IncludeProtection);
                Assert.Equal("IndentLevelProbe", style.Name);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
    }

    private static void AssertNoExternalInteropAssemblies()
    {
        string[] names = ["office", "Microsoft.Vbe.Interop", "Microsoft.Office.Interop.Excel", "ADODB"];
        var references = typeof(ExcelBatch).Assembly.GetReferencedAssemblies();
        var loaded = AppDomain.CurrentDomain.GetAssemblies();
        foreach (string name in names)
        {
            Assert.DoesNotContain(references, reference => reference.Name == name);
            Assert.DoesNotContain(loaded, assembly => assembly.GetName().Name == name);
            Assert.False(File.Exists(Path.Combine(AppContext.BaseDirectory, $"{name}.dll")));
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    [InlineData(30)]
    public void Execute_EmbeddedVbeDefinitions_RunPassesOptionalArgumentsWithoutRuntimeInteropDlls(int argumentCount)
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Vbe.VBProject? project = null;
            Vbe.VBComponents? components = null;
            Vbe.VBComponent? component = null;
            Vbe.CodeModule? code = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                project = ctx.Book.VBProject;
                components = project.VBComponents;
                component = components.Add(Vbe.vbext_ComponentType.vbext_ct_StdModule);
                component.Name = "EmbeddedRunProbe";
                code = component.CodeModule;
                string declaration = string.Join(", ", Enumerable.Range(1, 30).Select(index => $"Optional ByVal p{index} As Variant"));
                string statements = string.Join("\r\n", Enumerable.Range(1, 30).Select(index =>
                    $"""
                    If IsMissing(p{index}) Then
                        ThisWorkbook.Worksheets(1).Cells(1, {index}).Value = "missing"
                    Else
                        ThisWorkbook.Worksheets(1).Cells(1, {index}).Value = p{index}
                    End If
                    """));
                code.AddFromString($"Public Sub Store({declaration})\r\n{statements}\r\nEnd Sub");
                var args = Enumerable.Range(1, 30)
                    .Select(index => index <= argumentCount ? (object)$"value-{index}" : Type.Missing).ToArray();
                var originalSecurity = ctx.App.AutomationSecurity;
                var failures = new List<Exception>();
                try
                {
                    ctx.App.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityLow;
                    string procedure = $"'{ctx.Book.Name.Replace("'", "''", StringComparison.Ordinal)}'!EmbeddedRunProbe.Store";
                    ctx.App.Run(procedure,
                        args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9],
                        args[10], args[11], args[12], args[13], args[14], args[15], args[16], args[17], args[18], args[19],
                        args[20], args[21], args[22], args[23], args[24], args[25], args[26], args[27], args[28], args[29]);
                    sheets = ctx.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    range = sheet.Range["A1:AD1"];
                    object? storedValues = range.Value2;
                    var values = Assert.IsType<object[,]>(storedValues);
                    Assert.Equal(1, values.GetLength(0));
                    Assert.Equal(30, values.GetLength(1));
                    for (int index = 1; index <= 30; index++)
                    {
                        Assert.Equal(index <= argumentCount ? $"value-{index}" : "missing", values[1, index]);
                    }
                }
                catch (Exception ex)
                {
                    failures.Add(ex);
                }
                finally
                {
                    var restore = Record.Exception(() => ctx.App.AutomationSecurity = originalSecurity);
                    if (restore is not null) { failures.Add(restore); }
                }
                if (failures.Count > 0)
                {
                    throw new AggregateException("Embedded macro execution or security restoration failed.", failures);
                }
                Assert.Equal(originalSecurity, ctx.App.AutomationSecurity);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref code);
                ComUtilities.Release(ref component);
                ComUtilities.Release(ref components);
                ComUtilities.Release(ref project);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Fact]
    public void ExecuteAsync_MultipleOperations_ReusesExcelInstance()
    {
        // Arrange
        int operationCount = 0;

        // Act - Use batching for multiple operations
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
        var originalIdentity = ReadNativeIdentity(batch);
        var originalProcessId = batch.ExcelProcessId;
        SessionWorkbookAssertions.WriteMarker(batch, "initial");

        for (int i = 0; i < 5; i++)
        {
            batch.Execute((ctx, ct) =>
            {
                operationCount++;
                _output.WriteLine($"Batch operation {operationCount}");

                return operationCount;
            });
            Assert.Equal(originalIdentity, ReadNativeIdentity(batch));
            Assert.Equal(originalProcessId, batch.ExcelProcessId);
            Assert.Equal(i == 0 ? "initial" : $"operation-{i}",
                SessionWorkbookAssertions.ReadMarker(batch));
            SessionWorkbookAssertions.WriteMarker(batch, $"operation-{i + 1}");
        }

        // Assert
        Assert.Equal(5, operationCount);
        _output.WriteLine($"✓ Completed {operationCount} batch operations");
    }

    [Fact]
    public void Dispose_CleansUpComObjects_NoProcessLeak()
    {
        // Arrange
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
            SessionWorkbookAssertions.WriteMarker(batch, "disposed-workbook");
        }

        _owned.AssertAllExited();
    }

    [Fact]
    public void Save_PersistsChanges_ToWorkbook()
    {
        // Arrange
        string testValue = $"Test-{Guid.NewGuid():N}";

        // Act - Write and save
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
            SessionWorkbookAssertions.WriteMarker(batch, testValue);

            batch.Save();
        }

        _owned.AssertAllExited();

        // Verify - Read back the value in a new batch session
        string readValue;
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
            readValue = Assert.IsType<string>(SessionWorkbookAssertions.ReadMarker(batch));
        }

        // Assert
        Assert.Equal(testValue, readValue);
        _output.WriteLine($"✓ Value persisted correctly: {testValue}");
    }

    [Fact]
    public void Save_ReadOnlyWorkbook_RejectsEvenWhenAlreadySaved()
    {
        using var batch = ExcelSession.BeginReadOnlyValidation(
            _testFileCopy!, TimeSpan.FromSeconds(30));
        var originalMarker = SessionWorkbookAssertions.ReadMarker(batch);
        Assert.True(batch.Execute((context, _) => context.Book.ReadOnly));
        Assert.True(batch.Execute((context, _) => context.Book.Saved));

        var exception = Assert.Throws<InvalidOperationException>(() => batch.Save());

        Assert.Contains("read-only", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(originalMarker, SessionWorkbookAssertions.ReadMarker(batch));
        Assert.True(batch.Execute((context, _) => context.Book.ReadOnly));
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public void BeginBatch_ProtectedDetection_RequestsEditableAccess()
    {
        var syntheticProtectedPath = Path.Join(
            Path.GetTempPath(), $"batch-irm-access-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(syntheticProtectedPath);
        OleDataSpaceTestFile.Write(syntheticProtectedPath, "\tDRMDataSpace");
        Assert.True(FileAccessValidator.IsIrmProtected(syntheticProtectedPath));
        bool reachedOpen = false;
        ExcelBatch.BeforeWorkbookOpenHook = (path, _) =>
        {
            Assert.Equal(syntheticProtectedPath, path, ignoreCase: true);
            reachedOpen = true;
            // Replace only synthetic metadata with a valid fixture after detection.
            // This exercises the actual COM open options without enterprise credentials.
            File.Copy(_testFileCopy!, path, overwrite: true);
        };

        try
        {
            using var batch = ExcelSession.BeginBatch(
                show: true, operationTimeout: TimeSpan.FromSeconds(30), syntheticProtectedPath);
            Assert.True(reachedOpen);
            Assert.True(batch.Execute((context, _) => context.App.Visible));
            Assert.False(batch.Execute((context, _) => context.Book.ReadOnly));
            Assert.Equal(
                Path.GetFullPath(syntheticProtectedPath),
                batch.Execute((context, _) => context.Book.FullName),
                ignoreCase: true);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    [JapaneseLocaleFact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    [Trait("Locale", "ja-JP")]
    public void OpenAndSave_JapaneseTableDateFormat_PreservesTableColumnDxf()
    {
        string workbookPath = Path.Join(Path.GetTempPath(), $"table-dxf-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(workbookPath);
        CreateTableFormatWorkbook(workbookPath);
        string expectedFormat = ReadTableColumnFormatCode(workbookPath, "Date");
        Assert.Equal("yyyy/m/d", expectedFormat);

        using (var batch = ExcelSession.BeginBatch(workbookPath))
        {
            SessionWorkbookAssertions.WriteMarker(batch, "Japanese format retained");
            batch.Save();
        }

        Assert.Equal(expectedFormat, ReadTableColumnFormatCode(workbookPath, "Date"));
        using var reopened = ExcelSession.BeginBatch(workbookPath);
        Assert.Equal("Japanese format retained", SessionWorkbookAssertions.ReadMarker(reopened));
    }

    [Theory]
    [InlineData("yyyy/m/d")]
    [InlineData("#,##0.00")]
    public void OpenAndSave_TableColumnFormatAndEditedValue_SurviveReopening(string format)
    {
        string workbookPath = Path.Join(Path.GetTempPath(), $"table-format-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(workbookPath);
        CreateTableFormatWorkbook(workbookPath, format, localized: false);
        var expectedFormat = ReadTableColumnFormatCode(workbookPath, "Date");
        Assert.False(string.IsNullOrWhiteSpace(expectedFormat));

        using (var batch = ExcelSession.BeginBatch(workbookPath))
        {
            AssertTableValues(batch);
            SessionWorkbookAssertions.WriteMarker(batch, "First saved edit");
            batch.Save();
        }
        Assert.Equal(expectedFormat, ReadTableColumnFormatCode(workbookPath, "Date"));

        using (var reopened = ExcelSession.BeginBatch(workbookPath))
        {
            Assert.Equal("First saved edit", SessionWorkbookAssertions.ReadMarker(reopened));
            AssertTableValues(reopened);
            SessionWorkbookAssertions.WriteMarker(reopened, "Second saved edit");
            reopened.Save();
        }
        Assert.Equal(expectedFormat, ReadTableColumnFormatCode(workbookPath, "Date"));
        using var final = ExcelSession.BeginBatch(workbookPath);
        Assert.Equal("Second saved edit", SessionWorkbookAssertions.ReadMarker(final));
        AssertTableValues(final);

        static void AssertTableValues(IExcelBatch batch) => batch.Execute((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? values = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                values = sheet.Range["B2:B3"];
                var actual = Assert.IsType<object[,]>((object?)values.Value2);
                Assert.Equal(2, actual.GetLength(0));
                Assert.Equal(1, actual.GetLength(1));
                Assert.Equal(46000d, actual[1, 1]);
                Assert.Equal(46001d, actual[2, 1]);
            }
            finally
            {
                ComUtilities.Release(ref values);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public void Execute_DependentOperations_RetainCreatedWorksheetAndNamedRange()
    {
        const string sheetName = "TestData";
        const string namedRangeName = "TestRange";
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        var identity = ReadNativeIdentity(batch);
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Add();
                sheet.Name = sheetName;
                Assert.Equal(sheetName, sheet.Name);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        WithWorksheet(batch, sheetName, sheet =>
        {
            Excel.Range? values = null;
            Excel.Range? formula = null;
            try
            {
                values = sheet.Range["A1:B2"];
                values.Value2 = new object[,] { { "Header1", "Header2" }, { "Value1", 0 } };
                formula = sheet.Range["B2"];
                formula.Formula = "=LEN(A2)";
            }
            finally
            {
                ComUtilities.Release(ref formula);
                ComUtilities.Release(ref values);
            }
        });
        AssertWorkflowValues("Value1", 6);
        batch.Execute((context, _) =>
        {
            Excel.Names? names = null;
            Excel.Name? name = null;
            try
            {
                names = context.Book.Names;
                name = names.Add(namedRangeName, $"={sheetName}!$A$1:$B$2");
                Assert.Equal(namedRangeName, name.Name);
                Assert.Equal($"={sheetName}!$A$1:$B$2", (object?)name.RefersTo);
            }
            finally
            {
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        });
        WithWorksheet(batch, sheetName, sheet =>
        {
            Excel.Range? cell = null;
            try
            {
                cell = sheet.Range["A2"];
                cell.Value2 = "Modified";
            }
            finally { ComUtilities.Release(ref cell); }
        });
        AssertWorkflowValues("Modified", 8);
        batch.Execute((context, _) =>
        {
            Excel.Names? names = null;
            Excel.Name? name = null;
            Excel.Range? target = null;
            try
            {
                names = context.Book.Names;
                name = names.Item(namedRangeName);
                target = name.RefersToRange;
                Assert.Equal($"={sheetName}!$A$1:$B$2", (object?)name.RefersTo);
                Assert.Equal("$A$1:$B$2", target.Address);
                var values = Assert.IsType<object[,]>((object?)target.Value2);
                Assert.Equal("Modified", values[2, 1]);
                Assert.Equal(8d, values[2, 2]);
            }
            finally
            {
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        });
        Assert.Equal(identity, ReadNativeIdentity(batch));

        void AssertWorkflowValues(string expectedValue, double expectedLength)
        {
            WithWorksheet(batch, sheetName, sheet =>
            {
                Excel.Range? range = null;
                try
                {
                    range = sheet.Range["A1:B2"];
                    var values = Assert.IsType<object[,]>((object?)range.Value2);
                    Assert.Equal("Header1", values[1, 1]);
                    Assert.Equal("Header2", values[1, 2]);
                    Assert.Equal(expectedValue, values[2, 1]);
                    Assert.Equal(expectedLength, values[2, 2]);
                }
                finally { ComUtilities.Release(ref range); }
            });
        }
    }

    private static (int Window, IntPtr Workbook) ReadNativeIdentity(IExcelBatch batch) =>
        batch.Execute((context, _) =>
        {
            var workbook = Marshal.GetIUnknownForObject(context.Book);
            try { return (context.App.Hwnd, workbook); }
            finally { Marshal.Release(workbook); }
        });

    private static void WithWorksheet(IExcelBatch batch, object identifier, Action<Excel.Worksheet> action)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[identifier];
                action(sheet);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public async Task ParallelBatches_TwoConcurrentBatches_NoExcelProcessLeak()
    {
        // Arrange
        const int batchCount = 2;
        var testFileCopies = new List<string>();

        // Create fresh copies for parallel test
        for (int i = 0; i < batchCount; i++)
        {
            string copy = Path.Join(Path.GetTempPath(), $"batch-test-parallel-{i}-{Guid.NewGuid():N}.xlsx");
            _temporaryFiles.Add(copy);
            File.Copy(_staticTestFile!, copy, overwrite: true);
            testFileCopies.Add(copy);
        }

        var tasks = testFileCopies.Select((testFile, index) =>
        {
            return Task.Run(() =>
            {
                using var batch = ExcelSession.BeginBatch(testFile);
                SessionWorkbookAssertions.AssertIdentity(batch, testFile);
                var identity = ReadNativeIdentity(batch);

                // Perform multiple operations per batch
                for (int op = 0; op < 3; op++)
                {
                    WithWorksheet(batch, 1, sheet =>
                    {
                        Excel.Range? cell = null;
                        try
                        {
                            cell = sheet.Range[$"A{op + 1}"];
                            cell.Value2 = $"Batch{index}-Op{op}";
                            Assert.Equal($"Batch{index}-Op{op}", (object?)cell.Value2);
                        }
                        finally { ComUtilities.Release(ref cell); }
                    });
                    Assert.Equal(identity, ReadNativeIdentity(batch));
                }
                WithWorksheet(batch, 1, sheet =>
                {
                    Excel.Range? cells = null;
                    try
                    {
                        cells = sheet.Range["A1:A3"];
                        var values = Assert.IsType<object[,]>((object?)cells.Value2);
                        for (int row = 1; row <= 3; row++)
                        {
                            Assert.Equal($"Batch{index}-Op{row - 1}", values[row, 1]);
                        }
                    }
                    finally { ComUtilities.Release(ref cells); }
                });

                _output.WriteLine($"✓ Batch {index} completed");

                return (Index: index, ProcessId: batch.ExcelProcessId);
            });
        }).ToArray();

        // Wait for all batches to complete
        var results = await Task.WhenAll(tasks);

        Assert.Equal(batchCount, results.Length);
        Assert.Equal([0, 1], results.Select(result => result.Index).Order());
        Assert.All(results, result => Assert.NotNull(result.ProcessId));
        Assert.Equal(batchCount, results.Select(result => result.ProcessId).Distinct().Count());
        _output.WriteLine($"✓ All {batchCount} parallel batches completed");

        _owned.AssertAllExited();
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public void BeginBatch_IrmWorkbook_ShowFalse_FailsFastBeforeOpen()
    {
        string fakeIrmFile = Path.Join(Path.GetTempPath(), $"batch-irm-headless-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(fakeIrmFile);
        OleDataSpaceTestFile.Write(fakeIrmFile, "\tDRMDataSpace");
        var originalBytes = File.ReadAllBytes(fakeIrmFile);

        bool openAttempted = false;
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) => openAttempted = true;

        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelSession.BeginBatch(show: false, operationTimeout: TimeSpan.FromSeconds(15), fakeIrmFile));

            Assert.False(openAttempted);
            Assert.Contains("IRM/AIP-protected workbook", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("show=true", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(originalBytes, File.ReadAllBytes(fakeIrmFile));
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;

        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    public void BeginBatch_StartupFailureWithConfirmedExit_UntracksExactIdentity()
    {
        ExcelProcessIdentity? capturedIdentity = null;
        void CaptureIdentity(ExcelProcessIdentity identity)
        {
            capturedIdentity = identity;
            _startupIdentities.Add(identity);
        }

        SessionManager.ExcelProcessIdentityTracked += CaptureIdentity;
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) =>
            throw new InvalidOperationException("synthetic startup failure");
        ExcelBatch.FailedStartupExitConfirmationHook = _ => true;
        try
        {
            var exception = Assert.Throws<InvalidOperationException>(
                () => ExcelSession.BeginBatch(_testFileCopy!));

            Assert.Equal("synthetic startup failure", exception.Message);
            var identity = Assert.IsType<ExcelProcessIdentity>(capturedIdentity);
            Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());
        }
        finally
        {
            SessionManager.ExcelProcessIdentityTracked -= CaptureIdentity;
            ExcelBatch.BeforeWorkbookOpenHook = null;
            ExcelBatch.FailedStartupExitConfirmationHook = null;
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    public void BeginBatch_GenericExcel1004_PreservesOriginalDiagnostic()
    {
#pragma warning disable CA2201 // A real COMException is required to exercise Excel HRESULT classification.
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) =>
            throw new COMException("Synthetic non-lock Excel failure", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        try
        {
            var exception = Assert.Throws<COMException>(
                () => ExcelSession.BeginBatch(_testFileCopy!));

            Assert.Equal(unchecked((int)0x800A03EC), exception.HResult);
            Assert.Contains("Synthetic non-lock Excel failure", exception.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("locked by another process", exception.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    public void BeginBatch_StartupFailureWithUnconfirmedLiveIdentity_RetainsOwnership()
    {
        ExcelProcessIdentity? capturedIdentity = null;
        void CaptureIdentity(ExcelProcessIdentity identity)
        {
            capturedIdentity = identity;
            _startupIdentities.Add(identity);
        }

        SessionManager.ExcelProcessIdentityTracked += CaptureIdentity;
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) =>
            throw new InvalidOperationException("synthetic startup failure");
        ExcelBatch.FailedStartupExitConfirmationHook = _ => false;
        try
        {
            var exception = Assert.Throws<AggregateException>(
                () => ExcelSession.BeginBatch(_testFileCopy!));

            var identity = Assert.IsType<ExcelProcessIdentity>(capturedIdentity);
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains("remains tracked", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(
                exception.InnerExceptions,
                failure => failure.Message == "synthetic startup failure");
        }
        finally
        {
            SessionManager.ExcelProcessIdentityTracked -= CaptureIdentity;
            ExcelBatch.BeforeWorkbookOpenHook = null;
            ExcelBatch.FailedStartupExitConfirmationHook = null;
        }
    }

    [ConfiguredIrmFact]
    [Trait("RunType", "OnDemand")]
    public async Task BeginBatch_RealIrmWorkbook_CompletesStartupWithinBudget_WhenConfigured()
    {
        // Real IRM startup depends on interactive auth/enterprise policy and cannot run in CI.
        var irmTestFile = GetConfiguredIrmTestFilePath()
            ?? throw new InvalidOperationException("Configured IRM test fixture was unavailable after test discovery.");

        using var owned = new OwnedExcelProcessScope();

        var stopwatch = Stopwatch.StartNew();
        IExcelBatch? batch = null;
        var openTask = Task.Run(() =>
            ExcelSession.BeginBatch(show: true, operationTimeout: TimeSpan.FromSeconds(15), irmTestFile));
        var primary = await Record.ExceptionAsync(async () =>
        {
            batch = await openTask.WaitAsync(TimeSpan.FromSeconds(20));
            stopwatch.Stop();
            SessionWorkbookAssertions.AssertIdentity(batch, irmTestFile);
            Assert.True(batch.Execute((context, _) => context.App.Visible));
            Assert.True(stopwatch.Elapsed <= TimeSpan.FromSeconds(20));
            _output.WriteLine($"Opened IRM workbook in {stopwatch.Elapsed.TotalSeconds:F1}s");
        });
        var completion = await Record.ExceptionAsync(async () =>
        {
            if (batch is null && !openTask.IsFaulted && !openTask.IsCanceled)
            {
                batch = await openTask.WaitAsync(TimeSpan.FromSeconds(60));
            }
        });
        var disposal = Record.Exception(() => batch?.Dispose());
        var exit = Record.Exception(() => owned.AssertAllExited());
        var failures = new[] { primary, completion, disposal, exit }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
        {
            throw new AggregateException("Configured IRM startup or cleanup failed.", failures);
        }
    }

    private static void CreateTableFormatWorkbook(
        string workbookPath, string format = "yyyy/m/d", bool localized = true)
    {
        Excel.Application? excel = null;
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? workbook = null;
        Excel.Sheets? worksheets = null;
        Excel.Worksheet? worksheet = null;
        Excel.Range? sourceRange = null;
        Excel.ListObjects? listObjects = null;
        Excel.ListObject? table = null;
        Excel.ListColumns? listColumns = null;
        Excel.ListColumn? dateColumn = null;
        Excel.Range? dataBodyRange = null;
        Exception? primary = null;
        Exception? cleanup = null;

        try
        {
            primary = Record.Exception(() =>
            {
                excel = new Excel.Application { DisplayAlerts = false };
                workbooks = excel.Workbooks;
                workbook = workbooks.Add();
                worksheets = workbook.Worksheets;
                worksheet = (Excel.Worksheet)worksheets[1];
                sourceRange = worksheet.Range["A1:B3"];
                sourceRange.Value2 = new object[,] { { "Amount", "Date" }, { 1, 46000 }, { 2, 46001 } };
                listObjects = worksheet.ListObjects;
                table = listObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, sourceRange,
                    Type.Missing, Excel.XlYesNoGuess.xlYes);
                listColumns = table.ListColumns;
                dateColumn = listColumns["Date"];
                dataBodyRange = dateColumn.DataBodyRange;
                if (localized)
                    dataBodyRange.NumberFormatLocal = format;
                else
                    dataBodyRange.NumberFormat = format;
                workbook.SaveAs(workbookPath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            });
            cleanup = Record.Exception(() => CloseWorkbookAndQuitExcel(workbook, excel));
        }
        finally
        {
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref dateColumn);
            ComUtilities.Release(ref listColumns);
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sourceRange);
            ComUtilities.Release(ref worksheet);
            ComUtilities.Release(ref worksheets);
            ComUtilities.Release(ref workbook);
            ComUtilities.Release(ref workbooks);
            ComUtilities.Release(ref excel);
        }
        var failures = new[] { primary, cleanup }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
        {
            throw new AggregateException("Japanese native fixture creation or cleanup failed.", failures);
        }
    }

    private static string ReadTableColumnFormatCode(string workbookPath, string columnName)
    {
        using var archive = ZipFile.OpenRead(workbookPath);
        var tableEntry = archive.Entries.Single(entry => entry.FullName.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase));
        using var tableStream = tableEntry.Open();
        var tableDocument = XDocument.Load(tableStream);
        XNamespace spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var column = tableDocument.Descendants(spreadsheetNamespace + "tableColumn")
            .Single(element => element.Attribute("name")?.Value == columnName);
        int dxfId = int.Parse(column.Attribute("dataDxfId")!.Value, CultureInfo.InvariantCulture);

        var stylesEntry = archive.GetEntry("xl/styles.xml")
            ?? throw new InvalidOperationException("Workbook styles were not found.");
        using var stylesStream = stylesEntry.Open();
        var stylesDocument = XDocument.Load(stylesStream);
        return stylesDocument.Root!
            .Element(spreadsheetNamespace + "dxfs")!
            .Elements(spreadsheetNamespace + "dxf")
            .ElementAt(dxfId)
            .Element(spreadsheetNamespace + "numFmt")!
            .Attribute("formatCode")!
            .Value;
    }

    private static void CloseWorkbookAndQuitExcel(Excel.Workbook? workbook, Excel.Application? excel)
    {
        var failures = new List<Exception>();
        if (workbook != null)
        {
            var failure = Record.Exception(() => workbook.Close(false));
            if (failure is not null) { failures.Add(failure); }
        }

        if (excel != null)
        {
            var failure = Record.Exception(excel.Quit);
            if (failure is not null) { failures.Add(failure); }
        }
        if (failures.Count > 0)
        {
            throw new AggregateException("Native Excel fixture shutdown failed.", failures);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "FileLocking")]
    public void Constructor_FileLockedByAnotherProcess_ThrowsInvalidOperationException()
    {
        // Arrange - Create a separate test file for locking test
        var lockedTestFile = Path.Join(Path.GetTempPath(), $"batch-test-locked-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(lockedTestFile);
        File.Copy(_staticTestFile!, lockedTestFile, overwrite: true);
        var originalBytes = File.ReadAllBytes(lockedTestFile);

        {
            // Lock the file by opening with exclusive access (simulating Excel or another process)
            using var fileLock = new FileStream(
                lockedTestFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);

            // Act & Assert - Attempting to create ExcelBatch should fail immediately
            var ex = Assert.Throws<InvalidOperationException>(() =>
            {
                var batch = ExcelSession.BeginBatch(lockedTestFile);
                batch.Dispose();
            });

            // Verify error message is clear and actionable
            Assert.Contains("already open", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("close the file", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exclusive access", ex.Message, StringComparison.OrdinalIgnoreCase);

            _output.WriteLine($"✓ File locking detected successfully");
            _output.WriteLine($"Error message: {ex.Message}");
        }
        Assert.Equal(originalBytes, File.ReadAllBytes(lockedTestFile));
        using var recovered = ExcelSession.BeginBatch(lockedTestFile);
        SessionWorkbookAssertions.AssertIdentity(recovered, lockedTestFile);
        SessionWorkbookAssertions.WriteMarker(recovered, "recovered-after-file-lock");
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "FileLocking")]
    public void Constructor_FileLockedByAnotherProcess_DoesNotLeakExcelProcess()
    {
        var lockedTestFile = Path.Join(Path.GetTempPath(), $"batch-test-locked-leak-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(lockedTestFile);
        File.Copy(_staticTestFile!, lockedTestFile, overwrite: true);
        var originalBytes = File.ReadAllBytes(lockedTestFile);

        using var owned = new OwnedExcelProcessScope();

        {
            using var fileLock = new FileStream(
                lockedTestFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);

            var ex = Assert.Throws<InvalidOperationException>(() =>
            {
                using var batch = ExcelSession.BeginBatch(lockedTestFile);
            });

            Assert.Contains("already open", ex.Message, StringComparison.OrdinalIgnoreCase);

            Assert.Empty(owned.Identities);
            owned.AssertAllExited(expectProcess: false);
        }
        Assert.Equal(originalBytes, File.ReadAllBytes(lockedTestFile));
    }
}

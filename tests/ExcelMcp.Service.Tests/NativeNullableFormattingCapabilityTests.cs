using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "CellStyles")]
[Trait("Feature", "TableStyles")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class NativeNullableFormattingCapabilityTests(
    PersistentServiceWorkbookFixture fixture, ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnseededStyle_NativeGettersAndServicePreserveExactState(bool excludedBorder)
    {
        string sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string styleName = excludedBorder ? $"NativeDefault_{Guid.NewGuid():N}" : "Normal";
        bool created = false;
        Exception? primaryFailure = null;
        try
        {
            var observations = _fixture.ExecuteRawVerification((ctx, _) =>
            {
                Excel.Worksheet? sheet = null;
                Excel.Range? source = null;
                object? originalSheet = null;
                Excel.Styles? styles = null;
                Excel.Style? style = null;
                Excel.Font? font = null;
                Excel.Borders? borders = null;
                Excel.Border? border = null;
                try
                {
                    styles = ctx.Book.Styles;
                    if (excludedBorder)
                    {
                        sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                        Assert.NotNull(sheet);
                        source = sheet.Range["A1"];
                        source.ClearFormats();
                        Assert.Equal(0, Convert.ToInt32(source.IndentLevel, CultureInfo.InvariantCulture));
                        originalSheet = ctx.App.ActiveSheet;
                        sheet.Activate();
                        style = styles.Add(styleName, source);
                        created = true;
                        ComUtilities.Release(ref style);
                        style = styles[styleName];
                        style.IncludeAlignment = false;
                        style.IncludeProtection = false;
                        style.IncludeBorder = true;
                        borders = style.Borders;
                        border = borders[(Excel.XlBordersIndex)1];
                        border.LineStyle = Excel.XlLineStyle.xlDash;
                        border.Color = 0x563412;
                        Assert.False(style.IncludeAlignment);
                        Assert.False(style.IncludeProtection);
                        Assert.True(style.IncludeBorder);
                    }
                    else
                    {
                        style = styles[styleName];
                        Assert.True(style.BuiltIn);
                    }
                    font = style.Font;
                    Assert.Equal(typeof(Excel.Style).GUID, typeof(RangeCommands.IStyleIndentLevelVariant).GUID);
                    var variantStyle = (RangeCommands.IStyleIndentLevelVariant)style;
                    Assert.Same(style, variantStyle);
                    var variantIndent = Assert.IsType<DBNull>(variantStyle.IndentLevel);
                    output.WriteLine($"Style.IndentLevel typed native VARIANT binding: {variantIndent.GetType().FullName}; same RCW=true");
                    // Capability probe: record native VARIANTs separately from each typed getter.
                    dynamic native = style;
                    var reads = new (string Field, object? Native, Func<object?> Typed)[]
                    {
                        ("horizontalAlignment", native.HorizontalAlignment, () => style.HorizontalAlignment),
                        ("verticalAlignment", native.VerticalAlignment, () => style.VerticalAlignment),
                        ("wrapText", native.WrapText, () => style.WrapText),
                        ("shrinkToFit", native.ShrinkToFit, () => style.ShrinkToFit),
                        ("addIndent", native.AddIndent, () => style.AddIndent),
                        ("indentLevel", native.IndentLevel, () => style.IndentLevel),
                        ("orientation", native.Orientation, () => style.Orientation),
                        ("readingOrder", native.ReadingOrder, () => style.ReadingOrder),
                        ("locked", native.Locked, () => style.Locked),
                        ("formulaHidden", native.FormulaHidden, () => style.FormulaHidden),
                        ("font.themeFont", ((dynamic)font).ThemeFont, () => font.ThemeFont)
                    };
                    var result = new List<NativeGetterObservation>();
                    foreach (var read in reads)
                    {
                        object? typed = null;
                        var failure = Record.Exception(() => typed = read.Typed());
                        output.WriteLine($"{styleName}.{read.Field}: native type={read.Native?.GetType().FullName ?? "null"}, " +
                            $"native={read.Native ?? "<null>"}, typed={typed ?? "<null>"}, " +
                            $"failure={failure?.GetType().FullName ?? "none"}, " +
                            $"HRESULT=0x{failure?.HResult ?? 0:X8}, message={failure?.Message ?? "none"}");
                        if (read.Field == "indentLevel")
                        {
                            Assert.IsType<DBNull>(read.Native);
                            var typeMismatch = Assert.IsType<InvalidCastException>(failure);
                            Assert.Equal(unchecked((int)0x80004002), typeMismatch.HResult);
                            Assert.Contains("OleAut reported a type mismatch", typeMismatch.Message, StringComparison.Ordinal);
                        }
                        else
                        {
                            Assert.Null(failure);
                            Assert.NotNull(typed);
                            AssertNativeTypedEquivalent(read.Field, read.Native, typed);
                        }
                        result.Add(new(read.Field, read.Native));
                    }
                    return result;
                }
                finally
                {
                    try
                    {
                        if (originalSheet is Excel.Worksheet originalWorksheet) originalWorksheet.Activate();
                        else if (originalSheet is Excel.Chart originalChart) originalChart.Activate();
                    }
                    finally
                    {
                        ComUtilities.Release(ref originalSheet);
                        ComUtilities.Release(ref border);
                        ComUtilities.Release(ref borders);
                        ComUtilities.Release(ref font);
                        ComUtilities.Release(ref style);
                        ComUtilities.Release(ref styles);
                        ComUtilities.Release(ref source);
                        ComUtilities.Release(ref sheet);
                    }
                }
            });
            var response = _fixture.Send("workbook.get-cell-style", new { styleName });
            Assert.True(response.Success, response.ErrorMessage);
            var definition = ServiceCommandProxy.DeserializeResult<CellStyleResult>(response.Result!);
            Assert.Null(definition.Style.Format.IndentLevel);
            Assert.Contains("indentLevel", definition.Style.Format.MixedFields);
            using var document = JsonDocument.Parse(response.Result!);
            var styleResult = document.RootElement.GetProperty("style");
            var format = styleResult.GetProperty("format");
            var mixed = format.GetProperty("mixedFields").EnumerateArray().Select(field => field.GetString()).ToArray();
            foreach (var observation in observations)
            {
                var container = observation.Field == "font.themeFont" ? format.GetProperty("font") : format;
                string property = observation.Field == "font.themeFont" ? "themeFont" : observation.Field;
                if (observation.Native is null or DBNull)
                {
                    Assert.False(container.TryGetProperty(property, out _));
                    Assert.Contains(observation.Field, mixed);
                }
                else
                {
                    var value = container.GetProperty(property);
                    if (observation.Native is bool boolean)
                        Assert.Equal(boolean, value.GetBoolean());
                    else
                        Assert.Equal(Convert.ToInt32(observation.Native, CultureInfo.InvariantCulture), value.GetInt32());
                    Assert.DoesNotContain(observation.Field, mixed);
                }
            }
            if (excludedBorder)
            {
                Assert.False(styleResult.GetProperty("includeAlignment").GetBoolean());
                Assert.False(styleResult.GetProperty("includeProtection").GetBoolean());
                Assert.True(styleResult.GetProperty("includeBorder").GetBoolean());
                var border = Assert.Single(format.GetProperty("borders").EnumerateArray(),
                    item => item.GetProperty("edge").GetString() == "xlEdgeLeft");
                Assert.Equal(-4115, border.GetProperty("lineStyle").GetInt32());
                Assert.Equal("#123456", border.GetProperty("color").GetProperty("rgb").GetString());
            }
        }
        catch (Exception failure)
        {
            primaryFailure = failure;
        }
        finally
        {
            if (created)
            {
                try
                {
                    _fixture.ExecuteRawVerification((ctx, _) =>
                    {
                        Excel.Styles? styles = null;
                        Excel.Style? style = null;
                        try
                        {
                            styles = ctx.Book.Styles;
                            style = styles[styleName];
                            style.Delete();
                        }
                        finally
                        {
                            ComUtilities.Release(ref style);
                            ComUtilities.Release(ref styles);
                        }
                    });
                }
                catch (Exception cleanupFailure)
                {
                    primaryFailure = primaryFailure is null
                        ? cleanupFailure
                        : new AggregateException(primaryFailure, cleanupFailure);
                }
            }
        }
        if (primaryFailure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(primaryFailure).Throw();
    }

    private sealed record NativeGetterObservation(string Field, object? Native);

    [Fact]
    public void ExcludedStyle_GettersMatchTypedPiaAndExactServiceFields()
    {
        string sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string styleName = $"Nullable_{Guid.NewGuid():N}";
        string[] fields =
        [
            "horizontalAlignment", "verticalAlignment", "wrapText", "shrinkToFit",
            "addIndent", "indentLevel", "orientation", "readingOrder", "locked", "formulaHidden"
        ];
        var expected = new Dictionary<string, object>();
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? first = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                Assert.NotNull(sheet);
                first = sheet.Range["A1"];
                SetStyleSource(first);
            }
            finally
            {
                ComUtilities.Release(ref first);
                ComUtilities.Release(ref sheet);
            }
        });
        var created = _fixture.Send("workbook.create-cell-style", new
        {
            styleName,
            sourceSheetName = sheetName,
            sourceCellAddress = "A1"
        });
        Assert.True(created.Success, created.ErrorMessage);
        var excluded = _fixture.Send("workbook.update-cell-style", new
        {
            styleName,
            styleOptions = new { includeAlignment = false, includeProtection = false }
        });
        Assert.True(excluded.Success, excluded.ErrorMessage);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            try
            {
                styles = ctx.Book.Styles;
                style = styles[styleName];
                Assert.False(style.IncludeAlignment);
                Assert.False(style.IncludeProtection);
                // Compare native VARIANT values with supported typed getters on a single-cell style.
                dynamic native = style;
                object?[] values =
                [
                    native.HorizontalAlignment, native.VerticalAlignment, native.WrapText, native.ShrinkToFit,
                    native.AddIndent, native.IndentLevel, native.Orientation, native.ReadingOrder,
                    native.Locked, native.FormulaHidden
                ];
                object[] typed =
                [
                    style.HorizontalAlignment, style.VerticalAlignment,
                    style.WrapText, style.ShrinkToFit, style.AddIndent,
                    style.IndentLevel, style.Orientation, style.ReadingOrder,
                    style.Locked, style.FormulaHidden
                ];
                for (int index = 0; index < fields.Length; index++)
                {
                    AssertNativeTypedEquivalent($"Style.{fields[index]}", values[index], typed[index]);
                    expected.Add(fields[index], typed[index]);
                }
                Assert.Equal(Excel.XlHAlign.xlHAlignLeft, style.HorizontalAlignment);
                Assert.Equal(Excel.XlVAlign.xlVAlignTop, style.VerticalAlignment);
                Assert.True(style.WrapText);
                Assert.True(style.ShrinkToFit);
                Assert.True(style.AddIndent);
                Assert.Equal(2, style.IndentLevel);
                var variantStyle = (RangeCommands.IStyleIndentLevelVariant)style;
                Assert.Same(style, variantStyle);
                Assert.Equal(2, Assert.IsType<int>(variantStyle.IndentLevel));
                Assert.Equal(Excel.XlOrientation.xlHorizontal, style.Orientation);
                Assert.Equal(-5004, style.ReadingOrder);
                Assert.False(style.Locked);
                Assert.True(style.FormulaHidden);
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
        var response = _fixture.Send("workbook.get-cell-style", new { styleName });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        var definition = result.RootElement.GetProperty("style");
        Assert.False(definition.GetProperty("includeAlignment").GetBoolean());
        Assert.False(definition.GetProperty("includeProtection").GetBoolean());
        var format = definition.GetProperty("format");
        var mixedFields = format.GetProperty("mixedFields").EnumerateArray()
            .Select(field => field.GetString()).ToArray();
        foreach (string field in fields)
        {
            if (expected[field] is bool boolean)
                Assert.Equal(boolean, format.GetProperty(field).GetBoolean());
            else
                Assert.Equal(Convert.ToInt32(expected[field], CultureInfo.InvariantCulture), format.GetProperty(field).GetInt32());
            Assert.DoesNotContain(field, mixedFields);
        }
    }

    [Fact]
    public void MixedFontThemeFont_NativeNoneMatchesPiaAndExactPerCellThemeFonts()
    {
        string sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? first = null;
            Excel.Range? second = null;
            Excel.Range? mixed = null;
            Excel.Font? firstFont = null;
            Excel.Font? secondFont = null;
            Excel.Font? mixedFont = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                Assert.NotNull(sheet);
                first = sheet.Range["A1"];
                second = sheet.Range["B1"];
                firstFont = first.Font;
                secondFont = second.Font;
                firstFont.ThemeFont = Excel.XlThemeFont.xlThemeFontMajor;
                secondFont.ThemeFont = Excel.XlThemeFont.xlThemeFontMinor;
                Assert.Equal(Excel.XlThemeFont.xlThemeFontMajor, firstFont.ThemeFont);
                Assert.Equal(Excel.XlThemeFont.xlThemeFontMinor, secondFont.ThemeFont);
                mixed = sheet.Range["A1:B1"];
                mixedFont = mixed.Font;
                // Excel reports xlThemeFontNone for a range mixing major and minor theme fonts.
                object? native = ((dynamic)mixedFont).ThemeFont;
                Assert.Equal(0, Assert.IsType<int>(native));
                Assert.Equal(Excel.XlThemeFont.xlThemeFontNone, mixedFont.ThemeFont);
                AssertNativeTypedEquivalent("Font.ThemeFont (mixed range)", native, mixedFont.ThemeFont);
            }
            finally
            {
                ComUtilities.Release(ref mixedFont);
                ComUtilities.Release(ref secondFont);
                ComUtilities.Release(ref firstFont);
                ComUtilities.Release(ref mixed);
                ComUtilities.Release(ref second);
                ComUtilities.Release(ref first);
                ComUtilities.Release(ref sheet);
            }
        });
        var response = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = "A1:B1" });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        var cells = result.RootElement.GetProperty("cells").EnumerateArray().ToArray();
        Assert.Equal(2, cells.Length);
        Assert.Equal("$A$1", cells[0].GetProperty("address").GetString());
        Assert.Equal("$B$1", cells[1].GetProperty("address").GetString());
        Assert.Equal(1, cells[0].GetProperty("stored").GetProperty("font").GetProperty("themeFont").GetInt32());
        Assert.Equal(2, cells[1].GetProperty("stored").GetProperty("font").GetProperty("themeFont").GetInt32());
    }

    [Fact]
    public void DifferentialFontThemeFont_SupportedElementMatchesNativeAndServiceFields()
    {
        string styleName = $"NullableTheme_{Guid.NewGuid():N}";
        var created = _fixture.Send("workbook.create-table-style",
            new { styleName, sourceStyleName = "TableStyleMedium2" });
        Assert.True(created.Success, created.ErrorMessage);
        var cleared = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new { elements = new[] { new { elementType = "xlFirstHeaderCell", clear = true } } }
        });
        Assert.True(cleared.Success, cleared.ErrorMessage);
        using (var clearResult = JsonDocument.Parse(cleared.Result!))
        {
            var element = GetFirstHeaderCell(clearResult);
            Assert.False(element.GetProperty("hasFormat").GetBoolean());
            Assert.Equal(JsonValueKind.Null, element.GetProperty("font").ValueKind);
            Assert.Equal(JsonValueKind.Null, element.GetProperty("fill").ValueKind);
        }
        var formatted = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new { elements = new[] { new { elementType = "xlFirstHeaderCell", bold = true } } }
        });
        Assert.True(formatted.Success, formatted.ErrorMessage);
        int? expectedThemeFont = null;
        bool unavailable = false;
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.TableStyles? styles = null;
            Excel.TableStyle? style = null;
            Excel.TableStyleElements? elements = null;
            Excel.TableStyleElement? element = null;
            Excel.Font? font = null;
            try
            {
                styles = ctx.Book.TableStyles;
                style = styles[styleName];
                elements = style.TableStyleElements;
                element = elements.Item(Excel.XlTableStyleElementType.xlFirstHeaderCell);
                Assert.True(element.HasFormat);
                font = element.Font;
                Assert.True(Convert.ToBoolean(font.Bold, CultureInfo.InvariantCulture));
                // Inspect the element again after the Service write, rather than a stale font reference.
                object? native = ((dynamic)font).ThemeFont;
                var failure = Record.Exception(() => expectedThemeFont = (int)font.ThemeFont);
                if (failure is null)
                {
                    AssertNativeTypedEquivalent("Font.ThemeFont (differential element)", native, expectedThemeFont!.Value);
                }
                else
                {
                    output.WriteLine($"Font.ThemeFont (differential element): native={native ?? "<null>"}, " +
                        $"typed failure={failure.GetType().FullName}, HRESULT=0x{failure.HResult:X8}, message={failure.Message}");
                    Assert.True(native is null or DBNull);
                    Assert.True(failure is COMException or ArgumentException, failure.ToString());
                    Assert.Contains(failure.HResult, new[]
                    {
                        unchecked((int)0x800A03EC), unchecked((int)0x80070057)
                    });
                    unavailable = true;
                }
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref element);
                ComUtilities.Release(ref elements);
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
            }
        });
        var response = _fixture.Send("workbook.get-table-style", new { styleName });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        var header = GetFirstHeaderCell(result);
        Assert.True(header.GetProperty("hasFormat").GetBoolean());
        Assert.True(header.GetProperty("font").GetProperty("bold").GetBoolean());
        if (expectedThemeFont.HasValue)
        {
            Assert.Equal(expectedThemeFont.Value, header.GetProperty("font").GetProperty("themeFont").GetInt32());
            Assert.DoesNotContain("font.themeFont", header.GetProperty("unsetOrUnavailableFields").EnumerateArray()
                .Select(field => field.GetString()));
        }
        else
        {
            Assert.True(unavailable);
            Assert.Equal(JsonValueKind.Null, header.GetProperty("font").GetProperty("themeFont").ValueKind);
            var fields = header.GetProperty("unsetOrUnavailableFields").EnumerateArray()
                .Select(field => field.GetString()).ToArray();
            Assert.Contains("font.themeFont", fields);
            Assert.Contains("font.themeFont.unavailable", fields);
        }
        var updated = _fixture.Send("workbook.update-table-style", new
        {
            styleName,
            tableStyleOptions = new { elements = new[] { new { elementType = "xlFirstHeaderCell", themeFont = 1 } } }
        });
        Assert.True(updated.Success, updated.ErrorMessage);
        using var changed = JsonDocument.Parse(updated.Result!);
        var updatedHeader = GetFirstHeaderCell(changed);
        Assert.Equal(1, updatedHeader.GetProperty("font").GetProperty("themeFont").GetInt32());
        Assert.DoesNotContain("font.themeFont", updatedHeader.GetProperty("unsetOrUnavailableFields").EnumerateArray()
            .Select(field => field.GetString()));
    }

    private void AssertNativeTypedEquivalent(string member, object? native, object typed)
    {
        output.WriteLine($"{member}: native type={native?.GetType().FullName ?? "null"}, value={native ?? "<null>"}");
        Assert.NotNull(native);
        Assert.IsNotType<DBNull>(native);
        output.WriteLine($"{member}: typed value={typed}, type={typed.GetType().FullName}");
        if (typed is bool boolean)
            Assert.Equal(boolean, Convert.ToBoolean(native, CultureInfo.InvariantCulture));
        else
            Assert.Equal(Convert.ToInt32(typed, CultureInfo.InvariantCulture), Convert.ToInt32(native, CultureInfo.InvariantCulture));
    }

    private static JsonElement GetFirstHeaderCell(JsonDocument document) =>
        Assert.Single(document.RootElement.GetProperty("style").GetProperty("elements").EnumerateArray(),
            element => element.GetProperty("elementType").GetString() == "xlFirstHeaderCell");

    private static void SetStyleSource(Excel.Range cell)
    {
        cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        cell.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
        cell.WrapText = true;
        cell.ShrinkToFit = true;
        cell.AddIndent = true;
        cell.IndentLevel = 2;
        cell.Orientation = 0;
        cell.ReadingOrder = -5004;
        cell.Locked = false;
        cell.FormulaHidden = true;
    }
}

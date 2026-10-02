using System.Text.Json;
using System.IO.Compression;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "WorkbookTheme")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceWorkbookThemeTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void GetTheme_ReturnsAllNativeColorsAndFontScripts()
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            dynamic? nativeTheme = null;
            dynamic? fontScheme = null;
            try
            {
                nativeTheme = ((dynamic)context.Book).Theme;
                fontScheme = nativeTheme.ThemeFontScheme;
                string[] majorNames = ["Cambria", "Tahoma", "Yu Mincho"];
                string[] minorNames = ["Calibri", "Arial", "Yu Gothic"];
                for (int index = 1; index <= majorNames.Length; index++)
                {
                    dynamic? majorFont = null;
                    dynamic? minorFont = null;
                    try
                    {
                        majorFont = fontScheme.MajorFont(index);
                        minorFont = fontScheme.MinorFont(index);
                        majorFont.Name = majorNames[index - 1];
                        minorFont.Name = minorNames[index - 1];
                    }
                    finally
                    {
                        ComUtilities.Release(ref minorFont);
                        ComUtilities.Release(ref majorFont);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref fontScheme);
                ComUtilities.Release(ref nativeTheme);
            }
        });
        var response = _fixture.Send("workbook.get-theme", new { });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(12, result.RootElement.GetProperty("colors").GetArrayLength());
        Assert.Equal(3, result.RootElement.GetProperty("majorFonts").GetArrayLength());
        Assert.Equal(3, result.RootElement.GetProperty("minorFonts").GetArrayLength());
        Assert.All(result.RootElement.GetProperty("colors").EnumerateArray(),
            color => Assert.Matches("^#[0-9A-F]{6}$", color.GetProperty("rgb").GetString()!));
        _fixture.ExecuteRawVerification((context, _) => context.Book.Save());
        using var file = new FileStream(_fixture.WorkbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        var entry = archive.GetEntry("xl/theme/theme1.xml");
        Assert.NotNull(entry);
        using var stream = entry.Open();
        var theme = XDocument.Load(stream);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var scheme = Assert.Single(theme.Descendants(drawing + "clrScheme"));
        string[] slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
        string[] names = ["xlThemeColorDark1", "xlThemeColorLight1", "xlThemeColorDark2", "xlThemeColorLight2",
            "xlThemeColorAccent1", "xlThemeColorAccent2", "xlThemeColorAccent3", "xlThemeColorAccent4",
            "xlThemeColorAccent5", "xlThemeColorAccent6", "xlThemeColorHyperlink", "xlThemeColorFollowedHyperlink"];
        for (int index = 0; index < slots.Length; index++)
        {
            var definition = Assert.Single(scheme.Element(drawing + slots[index])!.Elements());
            string rgb = definition.Name.LocalName == "sysClr"
                ? definition.Attribute("lastClr")!.Value
                : definition.Attribute("val")!.Value;
            var actual = result.RootElement.GetProperty("colors")[index];
            Assert.Equal(index + 1, actual.GetProperty("index").GetInt32());
            Assert.Equal(names[index], actual.GetProperty("name").GetString());
            Assert.Equal("#" + rgb.ToUpperInvariant(), actual.GetProperty("rgb").GetString());
        }
        string[] scripts = ["Latin", "ComplexScript", "EastAsian"];
        string[] nativeScripts = ["latin", "cs", "ea"];
        foreach (var (property, element) in new[] { ("majorFonts", "majorFont"), ("minorFonts", "minorFont") })
        {
            var fonts = Assert.Single(theme.Descendants(drawing + element));
            Assert.Equal(3, nativeScripts.Select(script => fonts.Element(drawing + script)!.Attribute("typeface")!.Value)
                .Distinct(StringComparer.Ordinal).Count());
            for (int index = 0; index < scripts.Length; index++)
            {
                var actual = result.RootElement.GetProperty(property)[index];
                Assert.Equal(scripts[index], actual.GetProperty("script").GetString());
                Assert.Equal(fonts.Element(drawing + nativeScripts[index])!.Attribute("typeface")!.Value,
                    actual.GetProperty("name").GetString());
            }
        }
    }

    [Fact]
    public async Task ApplyTheme_ChangesNativeDefinitionsAndPersistsAfterReopen()
    {
        string path = _fixture.ExecuteRawVerification((ctx, _) =>
            Path.GetFullPath(Path.Combine(ctx.App.Path, "..", "Document Themes 16", "Archway.thmx")));
        Assert.True(File.Exists(path), "The native Office Archway theme is required for this capability test.");
        var prepared = _fixture.Send("workbook.apply-theme", new { themePath = path });
        Assert.True(prepared.Success, prepared.ErrorMessage);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            dynamic? nativeTheme = null;
            dynamic? colors = null;
            dynamic? accent = null;
            try
            {
                nativeTheme = ((dynamic)ctx.Book).Theme;
                colors = nativeTheme.ThemeColorScheme;
                accent = colors.Colors(5);
                accent.RGB = Convert.ToInt32(accent.RGB) ^ 0xFFFFFF;
            }
            finally
            {
                ComUtilities.Release(ref accent);
                ComUtilities.Release(ref colors);
                ComUtilities.Release(ref nativeTheme);
            }
        });
        var before = _fixture.Send("workbook.get-theme", new { });
        var response = _fixture.Send("workbook.apply-theme", new { themePath = path });
        using var result = JsonDocument.Parse(response.Result!);
        using var original = JsonDocument.Parse(before.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.NotEqual(original.RootElement.GetProperty("colors").GetRawText(), result.RootElement.GetProperty("colors").GetRawText());
        await _fixture.SaveAndReopenAsync();
        var read = _fixture.Send("workbook.get-theme", new { });
        using var reopened = JsonDocument.Parse(read.Result!);
        Assert.Equal(result.RootElement.GetProperty("colors").GetRawText(), reopened.RootElement.GetProperty("colors").GetRawText());
        Assert.Equal(result.RootElement.GetProperty("majorFonts").GetRawText(), reopened.RootElement.GetProperty("majorFonts").GetRawText());
        Assert.Equal(result.RootElement.GetProperty("minorFonts").GetRawText(), reopened.RootElement.GetProperty("minorFonts").GetRawText());
    }

    [Theory]
    [InlineData("relative")]
    [InlineData("extension")]
    [InlineData("missing")]
    [InlineData("malformed")]
    public async Task InvalidTheme_PreservesNativeDefinitions(string kind)
    {
        var before = _fixture.Send("workbook.get-theme", new { });
        string path = kind switch
        {
            "relative" => "relative.thmx",
            "extension" => _fixture.CreateInputFile(".txt", "not a theme"),
            "malformed" => _fixture.CreateInputFile(".thmx", "not a theme"),
            _ => Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.thmx")
        };
        var rejected = await _fixture.SendForFailureAsync("workbook.apply-theme", new { themePath = path });
        Assert.False(rejected.Success);
        var after = _fixture.Send("workbook.get-theme", new { });
        Assert.Equal(before.Result, after.Result);
    }

    [Fact]
    public void ApplyingTheme_UpdatesThemeSensitiveFillButPreservesFixedRgb()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        // Seed independently of the formatting writer under test.
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? themed = null;
            Excel.Interior? themedFill = null;
            Excel.Range? fixedCell = null;
            Excel.Interior? fixedFill = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                themed = sheet.Range["A1"];
                themedFill = themed.Interior;
                themedFill.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
                fixedCell = sheet.Range["B1"];
                fixedFill = fixedCell.Interior;
                fixedFill.Color = 0x563412;
            }
            finally
            {
                ComUtilities.Release(ref fixedFill);
                ComUtilities.Release(ref fixedCell);
                ComUtilities.Release(ref themedFill);
                ComUtilities.Release(ref themed);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        string path = _fixture.ExecuteRawVerification((ctx, _) =>
            Path.GetFullPath(Path.Combine(ctx.App.Path, "..", "Document Themes 16", "Archway.thmx")));
        Assert.True(File.Exists(path), "The native Office Archway theme is required for this capability test.");
        var applied = _fixture.Send("workbook.apply-theme", new { themePath = path });
        using var theme = JsonDocument.Parse(applied.Result!);
        var read = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = "A1:B1" });
        using var formatting = JsonDocument.Parse(read.Result!);
        var cells = formatting.RootElement.GetProperty("cells");
        Assert.Equal(theme.RootElement.GetProperty("colors")[4].GetProperty("rgb").GetString(),
            cells[0].GetProperty("stored").GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
        Assert.Equal("#123456", cells[1].GetProperty("stored").GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
    }
}

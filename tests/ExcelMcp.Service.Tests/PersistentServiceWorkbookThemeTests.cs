using System.Text.Json;
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
        var response = _fixture.Send("workbook.get-theme", new { });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(12, result.RootElement.GetProperty("colors").GetArrayLength());
        Assert.Equal(3, result.RootElement.GetProperty("majorFonts").GetArrayLength());
        Assert.Equal(3, result.RootElement.GetProperty("minorFonts").GetArrayLength());
        Assert.All(result.RootElement.GetProperty("colors").EnumerateArray(),
            color => Assert.Matches("^#[0-9A-F]{6}$", color.GetProperty("rgb").GetString()!));
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
        // The scalar formatting contract does not expose theme indices; prepare the native state directly.
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

using Sbroenne.ExcelMcp.ComInterop.Formatting;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "Formatting")]
[Trait("RequiresExcel", "false")]
public sealed class NumberFormatTranslatorTests
{
    [Theory]
    [InlineData("[Red]0.00", true)]
    [InlineData("[black]0.00", true)]
    [InlineData("[Blue]0.00", true)]
    [InlineData("[Cyan]0.00", true)]
    [InlineData("[Green]0.00", true)]
    [InlineData("[Magenta]0.00", true)]
    [InlineData("[White]0.00", true)]
    [InlineData("[Yellow]0.00", true)]
    [InlineData("[>=1.5]0.00;[Red]0.00", true)]
    [InlineData("0.00 \"[Red]\"", false)]
    [InlineData("0.00\\[Red]", false)]
    [InlineData("0.00*[Red]", false)]
    [InlineData("[Color3]0.00", false)]
    [InlineData("[Red", false)]
    [InlineData("\"unfinished [Red]", false)]
    public void ContainsNamedColor_DetectsOnlyActualDirectives(string format, bool expected) =>
        Assert.Equal(expected, NumberFormatTranslator.ContainsNamedColor(format));

    [Theory]
    [InlineData("General", "General")]
    [InlineData("[Red]0.00", "[Red]0,00")]
    [InlineData("[>=1.5]0.00;[Red]0.00", "[>=1,5]0,00;[Red]0,00")]
    [InlineData("mmm-yy", "mmm-yy")]
    [InlineData("yyyy-mm-dd hh:mm:ss", "yyyy-mm-dd hh:mm:ss")]
    [InlineData("$#,##0.00", "$#.##0,00")]
    [InlineData("0.00,,\"M\"", "0,00..\"M\"")]
    [InlineData("0.00 \"a,b.c\"\\m", "0,00 \"a,b.c\"\\m")]
    public void TranslateForChart_PreservesEnglishKeywordsAndUsesRegionalNumbers(string format, string expected)
    {
        var translator = CreateGermanTranslator();
        Assert.Equal(expected, translator.TranslateForChart(format));
        Assert.Equal(format, translator.TranslateFromLocale(expected));
    }

    [Theory]
    [InlineData("General", "Standard")]
    [InlineData("[Red]General", "[Red]Standard")]
    [InlineData("mmm-yy", "MMM-JJ")]
    [InlineData("dddd, mmmm d, yyyy", "TTTT, MMMM T, JJJJ")]
    public void TranslateToLocale_LocalApiKeywordsAndDateNames_UseNativeCodes(string format, string expected)
    {
        var translator = CreateGermanTranslator();
        Assert.Equal(expected, translator.TranslateToLocale(format));
        Assert.Equal(format, translator.TranslateFromLocale(expected));
    }

    [Theory]
    [InlineData(".", ",", "$#,##0.00,,\"M\"", "$#,##0.00,,\"M\"")]
    [InlineData(",", ".", "$#,##0.00,,\"M\"", "$#.##0,00..\"M\"")]
    [InlineData(",", ".", "0.0E+0", "0,0E+0")]
    [InlineData(",", ".", "\"1,000.00\" #,##0.00", "\"1,000.00\" #.##0,00")]
    [InlineData(",", ".", "[>=1.5]0.00;[Red]0.00", "[>=1,5]0,00;[Red]0,00")]
    [InlineData(",", ".", "[$-409]0.00", "[$-409]0,00")]
    public void TranslateToLocale_PreservesPrecisionScalingAndLiterals(
        string decimalSeparator, string thousandsSeparator, string format, string expected)
    {
        var translator = new NumberFormatTranslator(decimalSeparator, thousandsSeparator);
        Assert.Equal(expected, translator.TranslateToLocale(format));
        Assert.Equal(format, translator.TranslateFromLocale(expected));
    }

    [Theory]
    [InlineData("General")]
    [InlineData("Standard")]
    public void TranslateFromLocale_LocalGeneralName_ReturnsInvariantKeyword(string generalFormatName)
    {
        var translator = new NumberFormatTranslator(",", ".", generalFormatName);

        Assert.Equal("General", translator.TranslateFromLocale(generalFormatName));
    }

    [Theory]
    [InlineData("[Red]Standard", "[Red]General")]
    [InlineData("[>=1,5]Standard;Standard", "[>=1.5]General;General")]
    [InlineData("\"Standard\"0,00", "\"Standard\"0.00")]
    public void TranslateFromLocale_LocalGeneralKeyword_HandlesSectionsAndLiterals(string format, string expected)
    {
        var translator = new NumberFormatTranslator(",", ".", "Standard");

        Assert.Equal(expected, translator.TranslateFromLocale(format));
    }

    [Theory]
    [InlineData("0.00 \"Total\"", "0,00 \"Total\"")]
    [InlineData("0.00 \\T", "0,00 \\T")]
    [InlineData("0.00 _T", "0,00 _T")]
    [InlineData("0.00 *J", "0,00 *J")]
    [InlineData("0.00 [Rot]\"Jahr\"", "0,00 [Rot]\"Jahr\"")]
    public void TranslateToLocale_LocalDateLettersInLiterals_DoNotSkipInvariantNumberTranslation(
        string format,
        string expected)
    {
        var translator = CreateGermanTranslator();

        Assert.Equal(expected, translator.TranslateToLocale(format));
        Assert.Equal(format, translator.TranslateFromLocale(expected));
    }

    [Theory]
    [InlineData("TT.MM.JJJJ")]
    [InlineData("[Red]TT.MM.JJJJ")]
    public void TranslateToLocale_GenuineLocalizedDateTokens_AreUnchanged(string format)
    {
        var translator = CreateGermanTranslator();

        Assert.Equal(format, translator.TranslateToLocale(format));
    }

    private static NumberFormatTranslator CreateGermanTranslator() =>
        new(",", ".", "Standard", dayCode: "T", monthCode: "M", yearCode: "J");
}

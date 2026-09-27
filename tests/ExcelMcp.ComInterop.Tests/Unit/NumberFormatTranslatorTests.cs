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
    [InlineData(".", ",", "$#,##0.00,,\"M\"", "$#,##0.00,,\"M\"")]
    [InlineData(",", ".", "$#,##0.00,,\"M\"", "$#.##0,00..\"M\"")]
    [InlineData(",", ".", "0.0E+0", "0,0E+0")]
    [InlineData(",", ".", "\"1,000.00\" #,##0.00", "\"1,000.00\" #.##0,00")]
    public void TranslateToLocale_PreservesPrecisionScalingAndLiterals(
        string decimalSeparator, string thousandsSeparator, string format, string expected)
    {
        var translator = new NumberFormatTranslator(decimalSeparator, thousandsSeparator);
        Assert.Equal(expected, translator.TranslateToLocale(format));
        Assert.Equal(format, translator.TranslateFromLocale(expected));
    }
}

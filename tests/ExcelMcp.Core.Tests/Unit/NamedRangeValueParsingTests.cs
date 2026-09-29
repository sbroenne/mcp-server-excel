using System.Globalization;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Parameters")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class NamedRangeValueParsingTests
{
    [Fact]
    public void ParseWriteValue_DottedIdentifierUnderGermanCulture_RemainsText()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");

            var parsed = NamedRangeCommands.ParseWriteValue("2.0.13");

            Assert.Equal("2.0.13", Assert.IsType<string>(parsed));
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }
}

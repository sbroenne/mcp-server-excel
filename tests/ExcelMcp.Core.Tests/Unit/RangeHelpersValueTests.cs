using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class RangeHelpersValueTests
{
    [Theory]
    [InlineData("'=1+1", "'=1+1")]
    [InlineData("'2026-08-27", "'2026-08-27")]
    [InlineData("plain text", "'plain text")]
    [InlineData("2026-08-27", "'2026-08-27")]
    public void ConvertToCellValue_PreservesExistingTextPrefix(
        string input,
        string expected)
    {
        var prepared = TypedCellValueParser.Parse(input, 1, 1);

        var result = RangeHelpers.ConvertToCellValue(prepared, use1904DateSystem: false);

        Assert.Equal(expected, result);
    }
}

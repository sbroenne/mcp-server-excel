using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class PivotCurrencyLiteralTests
{
    [Theory]
    [InlineData("$#,##0.00", "\\$#,##0.00")]
    [InlineData("\"$\"#,##0.00", "\\$#,##0.00")]
    [InlineData("\\$#,##0.00", "\\$#,##0.00")]
    [InlineData("[$$-409]#,##0.00", "[$$-409]#,##0.00")]
    [InlineData("[Red]$0.00;[Blue]-$0.00", "[Red]\\$0.00;[Blue]-\\$0.00")]
    [InlineData("0.00 \"USD $\"", "0.00 \"USD $\"")]
    [InlineData("0.00\\\"$\"text\"", "0.00\\\"\\$\"text\"")]
    [InlineData("_$0.00", "_$0.00")]
    [InlineData("*$0.00", "*$0.00")]
    [InlineData("[>=100]$0.00", "[>=100]\\$0.00")]
    [InlineData("0.00%", "0.00%")]
    [InlineData("yyyy-mm-dd", "yyyy-mm-dd")]
    [InlineData("", "")]
    public void PreserveCurrencyLiterals_ProtectsOnlyUnescapedDollars(string input, string expected)
    {
        var helpers = typeof(PivotTableCommands).Assembly.GetType(
            "Sbroenne.ExcelMcp.Core.Commands.NumberFormatLiterals", throwOnError: true)!;
        var method = helpers.GetMethod("PreserveCurrencyLiterals", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        Assert.Equal(expected, Assert.IsType<string>(method.Invoke(null, [input])));
    }
}

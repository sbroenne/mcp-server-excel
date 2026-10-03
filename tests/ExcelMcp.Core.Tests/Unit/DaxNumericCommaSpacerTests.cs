using Sbroenne.ExcelMcp.Core.DataModel;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public class DaxNumericCommaSpacerTests
{
    [Theory]
    [InlineData("IF(1=1, 1.5, 0)", "IF(1=1 , 1.5 , 0)")]
    [InlineData("DATEADD(T[Date], -1, MONTH)", "DATEADD(T[Date], -1 , MONTH)")]
    [InlineData("ROUND(1.25,1)", "ROUND(1.25 , 1)")]
    [InlineData("DIVIDE(SUM(T[A]),4)", "DIVIDE(SUM(T[A]), 4)")]
    [InlineData("MAX(.5,2)", "MAX(.5 , 2)")]
    [InlineData("MAX(1,.5)", "MAX(1 , .5)")]
    [InlineData("IF(x,1,0)", "IF(x, 1 , 0)")]
    public void AddSpaces_NumberTouchesComma_SeparatesThem(string input, string expected)
    {
        Assert.Equal(expected, DaxNumericCommaSpacer.AddSpaces(input));
    }

    [Theory]
    [InlineData("SUM(T[Amount])")]
    [InlineData("DIVIDE(SUM(T[A]), SUM(T[B]))")]
    [InlineData("IF(1=1 , 1.5 , 0)")]
    [InlineData("VAR x1 = 4 RETURN MAX(x1, x2)")]
    [InlineData("SUM(Table2[Col3])")]
    [InlineData("IF(\"1,2\" = \"1,2\", TRUE(), FALSE())")]
    [InlineData("IF(\"say \"\"1,2\"\"\" = \"x\", TRUE(), FALSE())")]
    [InlineData("SUM('Sales 2024,1'[Amount])")]
    [InlineData("SUM('It''s 1,2'[Amount])")]
    [InlineData("SUM(T[Col 1,2])")]
    [InlineData("SUM(T[A]]1,2])")]
    [InlineData("SUM(T[A]) // note 1,2")]
    [InlineData("SUM(T[A]) -- note 1,2")]
    [InlineData("SUM(T[A]) /* note 1,2 */")]
    [InlineData("")]
    public void AddSpaces_NoNumberTouchingArgumentComma_LeavesFormulaUnchanged(string input)
    {
        Assert.Equal(input, DaxNumericCommaSpacer.AddSpaces(input));
    }

    [Fact]
    public void AddSpaces_CommentEndsAtLineBreak_SpacesFollowingCode()
    {
        const string input = "// 1,2\nIF(1=1, 2, 3)";

        Assert.Equal("// 1,2\nIF(1=1 , 2 , 3)", DaxNumericCommaSpacer.AddSpaces(input));
    }

    [Fact]
    public void AddSpaces_StringBeforeNumericArguments_SpacesOnlyCodeCommas()
    {
        const string input = "IF(\"North, \"\"South\"\"\" = \"North, \"\"South\"\"\", 1.5, 0)";

        Assert.Equal(
            "IF(\"North, \"\"South\"\"\" = \"North, \"\"South\"\"\", 1.5 , 0)",
            DaxNumericCommaSpacer.AddSpaces(input));
    }

    [Fact]
    public void AddSpaces_UnterminatedString_LeavesRemainderUnchanged()
    {
        const string input = "IF(1, \"open 1,2";

        Assert.Equal("IF(1 , \"open 1,2", DaxNumericCommaSpacer.AddSpaces(input));
    }

    [Theory]
    [InlineData(",", true)]
    [InlineData(".", false)]
    [InlineData("٫", false)]
    [InlineData(null, false)]
    public void IsNeeded_DependsOnWindowsDecimalMark(string? decimalSeparator, bool expected)
    {
        Assert.Equal(expected, DaxNumericCommaSpacer.IsNeededFor(decimalSeparator));
    }
}

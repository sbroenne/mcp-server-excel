using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotReportingValidationTests
{
    [Fact]
    public void SubtotalFunctions_RejectNullBeforeBatch() =>
        Assert.Throws<ArgumentNullException>(() =>
            new PivotTableCommands().SetSubtotalFunctions(null!, "Pivot", "Region", null!));

    [Theory]
    [InlineData(PivotSubtotalFunction.Sum, PivotSubtotalFunction.Sum)]
    [InlineData(PivotSubtotalFunction.Automatic, PivotSubtotalFunction.Count)]
    [InlineData((PivotSubtotalFunction)0, PivotSubtotalFunction.Sum)]
    [InlineData((PivotSubtotalFunction)13, PivotSubtotalFunction.Sum)]
    public void SubtotalFunctions_RejectInvalidCombinationBeforeBatch(PivotSubtotalFunction first, PivotSubtotalFunction second)
    {
        var error = Assert.Throws<ArgumentException>(() =>
            new PivotTableCommands().SetSubtotalFunctions(null!, "Pivot", "Region", [first, second]));
        Assert.Equal("functions", error.ParamName);
    }

    [Fact]
    public void Sort_RejectUnknownDirectionBeforeBatch() =>
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PivotTableCommands().SortField(null!, "Pivot", "Region", (SortDirection)int.MaxValue));

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void CalculatedItem_RejectEmptyNameAndFormulaBeforeBatch(string value)
    {
        var commands = new PivotTableCommands();
        Assert.Throws<ArgumentException>(() => commands.CreateCalculatedItem(null!, "Pivot", "Region", value, "=North"));
        Assert.Throws<ArgumentException>(() => commands.CreateCalculatedItem(null!, "Pivot", "Region", "Combined", value));
        Assert.Throws<ArgumentException>(() => commands.DeleteCalculatedItem(null!, "Pivot", "Region", value));
    }
}

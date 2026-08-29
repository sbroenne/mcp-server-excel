using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Fast")]
public class TableFormulaReferenceTests
{
    [Fact]
    public void HasSortSensitiveReference_A1LookingStrings_ReturnsFalse()
    {
        const string formula = "=HYPERLINK(\"#A1\",\"Jump to \"\"B2\"\"\")";

        bool result = TableCommands.HasSortSensitiveReference(
            formula,
            formulaRow: 2,
            firstTableColumn: 6,
            lastTableColumn: 7);

        Assert.False(result);
    }

    [Fact]
    public void HasSortSensitiveReference_TrueReferenceAfterA1LookingString_ReturnsTrue()
    {
        const string formula = "=IF(\"A1\"=\"B2\",$F$2,F2)";

        bool result = TableCommands.HasSortSensitiveReference(
            formula,
            formulaRow: 2,
            firstTableColumn: 6,
            lastTableColumn: 7);

        Assert.True(result);
    }
}

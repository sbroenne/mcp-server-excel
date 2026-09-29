using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public class TableNameValidationTests
{
    // Service calls cannot verify that Core validates the name before using the batch.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_WithMissingTableName_ThrowsBeforeUsingBatch(string? tableName)
    {
        var commands = new TableCommands();
        var exception = Assert.ThrowsAny<ArgumentException>(() =>
            commands.Create(
                null!,
                "Data",
                tableName!,
                "A1:B2",
                true,
                "TableStyleLight1"));

        Assert.Equal("tableName", exception.ParamName);
    }
}

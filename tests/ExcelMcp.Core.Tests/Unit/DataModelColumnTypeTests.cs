using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public class DataModelColumnTypeTests
{
    [Theory]
    [InlineData(5, "DOUBLE")]
    [InlineData(7, "DATE")]
    [InlineData(11, "BOOLEAN")]
    [InlineData(20, "BIGINT")]
    [InlineData(130, "WCHAR")]
    [InlineData(0, "Unknown (0)")]
    [InlineData(8, "Unknown (8)")]
    [InlineData(-5, "Unknown (-5)")]
    [InlineData(999, "Unknown (999)")]
    public void GetModelColumnDataTypeName_UsesObservedOleDbCodesOrPreservesUnknownCode(
        int code, string expected)
    {
        Assert.Equal(expected, DataModelCommands.GetModelColumnDataTypeName(code));
    }
}

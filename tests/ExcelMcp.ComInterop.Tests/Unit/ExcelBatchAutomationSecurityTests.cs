using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelBatchAutomationSecurityTests
{
    [Theory]
    [InlineData(false, 3)]
    [InlineData(true, 1)]
    public void SelectAutomationSecurity_RequiresExplicitApproval(bool approved, int expected)
    {
        Assert.Equal(expected, ExcelBatch.SelectAutomationSecurity(approved));
    }
}

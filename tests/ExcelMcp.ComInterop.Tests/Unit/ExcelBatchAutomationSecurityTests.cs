using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelBatchAutomationSecurityTests
{
    [Fact]
    public void SelectAutomationSecurity_ReadOnlyMacroWorkbook_ForcesMacrosDisabled()
    {
        var security = ExcelBatch.SelectAutomationSecurity(
            createsMacroEnabledWorkbook: false,
            openReadOnly: true,
            workbookPaths: [@"C:\Workbooks\Untrusted.xlsm"]);

        Assert.Equal(3, security);
    }

    [Fact]
    public void SelectAutomationSecurity_RegularMacroWorkbook_PreservesMacroOperations()
    {
        var security = ExcelBatch.SelectAutomationSecurity(
            createsMacroEnabledWorkbook: false,
            openReadOnly: false,
            workbookPaths: [@"C:\Workbooks\Trusted.xlsm"]);

        Assert.Equal(1, security);
    }
}

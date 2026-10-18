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
    public void SelectAutomationSecurity_RegularMacroWorkbook_FollowsUserMacroSettings()
    {
        var security = ExcelBatch.SelectAutomationSecurity(
            createsMacroEnabledWorkbook: false,
            openReadOnly: false,
            workbookPaths: [@"C:\Workbooks\Untrusted.xlsm"]);

        Assert.Equal(2, security);
    }

    [Fact]
    public void SelectAutomationSecurity_CreatedMacroWorkbook_FollowsUserMacroSettings()
    {
        var security = ExcelBatch.SelectAutomationSecurity(
            createsMacroEnabledWorkbook: true,
            openReadOnly: false,
            workbookPaths: [@"C:\Workbooks\New.xlsm"]);

        Assert.Equal(2, security);
    }

    [Fact]
    public void SelectAutomationSecurity_RegularWorkbook_ForcesMacrosDisabled()
    {
        var security = ExcelBatch.SelectAutomationSecurity(
            createsMacroEnabledWorkbook: false,
            openReadOnly: false,
            workbookPaths: [@"C:\Workbooks\Data.xlsx"]);

        Assert.Equal(3, security);
    }
}

using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SaveBusyDiagnosticTests
{
    [Fact]
    public void SaveBusy_DoesNotInventAnExternalLock_AndExplainsUnsavedState()
    {
        var message = ExcelShutdownService.CreateSaveFailureMessage(
            unchecked((int)0x800AC472), "fixture.xlsx", "Excel declined automation.");

        Assert.Contains("Excel is busy", message, StringComparison.Ordinal);
        Assert.Contains("not been saved", message, StringComparison.Ordinal);
        Assert.Contains("open workbook", message, StringComparison.Ordinal);
        Assert.DoesNotContain("another user", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("dialog", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("locked", message, StringComparison.OrdinalIgnoreCase);
    }
}

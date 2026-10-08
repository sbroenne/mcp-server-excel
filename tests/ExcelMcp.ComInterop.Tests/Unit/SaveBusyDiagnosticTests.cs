using System.Runtime.InteropServices;
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
    public void FinalSaveBusy_PreservesTypedRefusalAndInnerComHResult()
    {
        var comFailure = Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800AC472)));
        var failure = ExcelShutdownService.CreateSaveFailureException(comFailure, "fixture.xlsx");

        var busy = Assert.IsType<ExcelBusyException>(failure);
        Assert.Same(comFailure, busy.InnerException);
        Assert.Equal(unchecked((int)0x800AC472), busy.InnerException.HResult);
        Assert.Contains("not been saved", busy.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(unchecked((int)0x800A03EC))]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED)]
    public void OtherSaveFailures_RetainTheirExistingCategoryAndComCause(int hresult)
    {
        var comFailure = Assert.IsType<COMException>(Marshal.GetExceptionForHR(hresult));
        var failure = ExcelShutdownService.CreateSaveFailureException(comFailure, "fixture.xlsx");

        Assert.IsType<InvalidOperationException>(failure);
        Assert.Same(comFailure, failure.InnerException);
    }

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

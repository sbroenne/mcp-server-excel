using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

#pragma warning disable CA2201 // Tests fabricate COM failures to verify retry behavior.

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public class ExcelShutdownServiceTests
{
    [Fact]
    public void CloseWorkbookWithRetry_RetriesTransientFailures()
    {
        int attempts = 0;

        ExcelShutdownService.CloseWorkbookWithRetry(() =>
        {
            attempts++;
            if (attempts < 3)
            {
                throw new COMException(
                    "Excel is busy.",
                    ResiliencePipelines.RPC_E_CALL_REJECTED);
            }
        });

        Assert.Equal(3, attempts);
    }

    [Fact]
    public void CloseWorkbookWithRetry_PropagatesNonTransientFailure()
    {
        var expected = new COMException("Close failed.", unchecked((int)0x800A03EC));

        var actual = Assert.Throws<COMException>(() =>
            ExcelShutdownService.CloseWorkbookWithRetry(() => throw expected));

        Assert.Same(expected, actual);
    }

    [Fact]
    public void CloseWorkbookWithRetry_WhenTransientRetriesAreExhausted_Propagates()
    {
        int attempts = 0;

        Assert.Throws<COMException>(() =>
            ExcelShutdownService.CloseWorkbookWithRetry(() =>
            {
                attempts++;
                throw new COMException(
                    "Excel remains busy.",
                    ResiliencePipelines.RPC_E_SERVERCALL_RETRYLATER);
            }));

        Assert.Equal(3, attempts);
    }
}

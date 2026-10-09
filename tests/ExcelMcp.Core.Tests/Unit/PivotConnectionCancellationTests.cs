using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "false")]
public sealed class PivotConnectionCancellationTests
{
    [Theory]
    [InlineData("success")]
    [InlineData("provider-error")]
    [InlineData("cancelled")]
    public void NativeConnectionChange_RegistersCancellationAndClearsItForRecovery(string outcome)
    {
        using var lifetime = new CancellationTokenSource();
        var entered = false;
        Action nativeCall = () =>
        {
            entered = true;
            Assert.Equal(lifetime.Token, PendingToken());
            Assert.NotEqual(0, MessagePending());
            if (outcome == "cancelled")
            {
                lifetime.Cancel();
                Assert.Equal(0, MessagePending());
            }
            if (outcome != "success")
            {
#pragma warning disable CA2201 // Simulated provider abort at the internal COM-call boundary.
                throw new COMException("Synthetic provider failure",
                    outcome == "cancelled" ? unchecked((int)0x80010002) : unchecked((int)0x800A03EC));
#pragma warning restore CA2201
            }
        };

        var failure = Record.Exception(() => Invoke(nativeCall, lifetime.Token));
        Assert.True(entered);
        if (outcome == "success")
            Assert.Null(failure);
        else
            Assert.IsType<COMException>(Assert.IsType<TargetInvocationException>(failure).InnerException);
        Assert.Equal(CancellationToken.None, PendingToken());
        Assert.NotEqual(0, MessagePending());

        var recovered = false;
        Invoke(() => recovered = true, CancellationToken.None);
        Assert.True(recovered);
        Assert.Equal(CancellationToken.None, PendingToken());
    }

    [Fact]
    public void NativeConnectionChange_CancelledBeforeCall_DoesNotEnterProvider()
    {
        using var lifetime = new CancellationTokenSource();
        lifetime.Cancel();
        var entered = false;
        var failure = Assert.Throws<TargetInvocationException>(() =>
            Invoke(() => entered = true, lifetime.Token));
        Assert.IsType<OperationCanceledException>(failure.InnerException);
        Assert.False(entered);
        Assert.Equal(CancellationToken.None, PendingToken());
    }

    [Fact]
    public void NativeConnectionChange_ProviderReturnsAfterCancellation_DoesNotReportSuccess()
    {
        using var lifetime = new CancellationTokenSource();
        var failure = Assert.Throws<TargetInvocationException>(() => Invoke(lifetime.Cancel, lifetime.Token));
        var cancellation = Assert.IsType<OperationCanceledException>(failure.InnerException);
        Assert.Equal(lifetime.Token, cancellation.CancellationToken);
        Assert.Equal(CancellationToken.None, PendingToken());
        Assert.NotEqual(0, MessagePending());
    }

    private static void Invoke(Action action, CancellationToken token)
    {
        var method = typeof(PivotTableCommands).GetMethod("ExecuteConnectionChange",
            BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        method.Invoke(null, [action, token]);
    }

    private static CancellationToken PendingToken()
    {
        var field = typeof(OleMessageFilter).GetField("_pendingCancellationToken",
            BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(field);
        return Assert.IsType<CancellationToken>(field.GetValue(null));
    }

    private static int MessagePending()
    {
        var contract = typeof(OleMessageFilter).Assembly.GetType("Sbroenne.ExcelMcp.ComInterop.IOleMessageFilter");
        Assert.NotNull(contract);
        var method = contract.GetMethod("MessagePending");
        Assert.NotNull(method);
        return Assert.IsType<int>(method.Invoke(new OleMessageFilter(), [nint.Zero, 0, 1]));
    }
}

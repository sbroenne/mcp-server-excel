using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ConnectionRefreshWaitTests
{
    [Fact]
    public void RefreshWait_CancelledBeforeIdleRead_DoesNotReturnSuccess()
    {
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        var cancelCalled = false;
        var error = Assert.Throws<TargetInvocationException>(() =>
            GetRefreshWaitMethod().Invoke(null,
            [
                (Func<bool>)(() => false),
                (Action)(() => cancelCalled = true),
                cancelled.Token
            ]));
        Assert.IsType<OperationCanceledException>(error.InnerException);
        Assert.True(cancelCalled);
    }

    [Fact]
    public void RefreshWait_StatusReadFailure_IsNotCompletion()
    {
#pragma warning disable CA2201 // Synthetic native status failure.
        var native = new COMException("Could not inspect refresh", unchecked((int)0x800AC472));
#pragma warning restore CA2201
        var error = Assert.Throws<TargetInvocationException>(() =>
            GetRefreshWaitMethod().Invoke(null,
            [
                (Func<bool>)(() => throw native),
                (Action)(() => throw new InvalidOperationException("Must not cancel automatically")),
                CancellationToken.None
            ]));
        Assert.Same(native, error.InnerException);
    }

    [Fact]
    public void RefreshWait_CancellationFailure_PreservesCancellationAndProviderCause()
    {
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
#pragma warning disable CA2201 // Synthetic provider cancellation failure.
        var native = new COMException("Provider did not cancel", unchecked((int)0x80004005));
#pragma warning restore CA2201
        var error = Assert.Throws<TargetInvocationException>(() =>
            GetRefreshWaitMethod().Invoke(null,
            [
                (Func<bool>)(() => true),
                (Action)(() => throw native),
                cancelled.Token
            ]));
        var cancellation = Assert.IsType<OperationCanceledException>(error.InnerException);
        Assert.Equal(cancelled.Token, cancellation.CancellationToken);
        Assert.Contains("may still be running", cancellation.Message, StringComparison.Ordinal);
        var causes = Assert.IsType<AggregateException>(cancellation.InnerException);
        Assert.Contains(native, causes.InnerExceptions);
    }

    [Fact]
    public void RefreshWait_WhenCancellationRequested_InvokesCancelActionAndThrows()
    {
        MethodInfo waitMethod = GetRefreshWaitMethod();

        bool cancelCalled = false;
        using var cts = new CancellationTokenSource();

        var cancellationThread = new Thread(() =>
        {
            Thread.Sleep(50);
            cts.Cancel();
        });
        cancellationThread.Start();

        try
        {
            var exception = Assert.Throws<TargetInvocationException>(() =>
                waitMethod.Invoke(null,
                [
                    (Func<bool>)(() => true),
                    (Action)(() => cancelCalled = true),
                    cts.Token
                ]));

            Assert.IsType<OperationCanceledException>(exception.InnerException);
            Assert.True(cancelCalled);
        }
        finally
        {
            cancellationThread.Join();
        }
    }

    [Fact]
    public void RefreshWait_WhenRefreshCompletes_DoesNotInvokeCancelAction()
    {
        MethodInfo waitMethod = GetRefreshWaitMethod();

        int pollCount = 0;
        bool cancelCalled = false;

        waitMethod.Invoke(null,
        [
            (Func<bool>)(() => Interlocked.Increment(ref pollCount) == 1),
            (Action)(() => cancelCalled = true),
            CancellationToken.None
        ]);

        Assert.True(pollCount >= 2);
        Assert.False(cancelCalled);
    }

    private static MethodInfo GetRefreshWaitMethod()
    {
        var waitMethod = typeof(ConnectionCommands).GetMethod(
            "WaitForConnectionRefreshCompletion",
            BindingFlags.NonPublic | BindingFlags.Static);

        return waitMethod ?? throw new InvalidOperationException(
            "Expected private method WaitForConnectionRefreshCompletion was not found.");
    }
}

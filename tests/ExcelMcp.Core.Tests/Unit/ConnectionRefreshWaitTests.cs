using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ConnectionRefreshWaitTests
{
    [Fact]
    public void OlapCapability_OdbcConnection_DoesNotReadOleDbMetadata()
    {
        var connection = ExcelInterfaceProxy.Create<Excel.WorkbookConnection>((member, _) =>
            member.Name == "get_Type"
                ? Enum.ToObject(member.ReturnType, (int)Excel.XlConnectionType.xlConnectionTypeODBC)
                : throw new InvalidOperationException($"Unexpected connection read: {member.Name}"));
        var method = typeof(ConnectionCommands).GetMethod("IsOlapConnection", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        Assert.False(Assert.IsType<bool>(method.Invoke(null, [(object)connection])));
    }

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

    internal static MethodInfo GetRefreshWaitMethod()
    {
        var waitMethod = typeof(ConnectionCommands).GetMethod(
            "WaitForConnectionRefreshCompletion",
            BindingFlags.NonPublic | BindingFlags.Static);

        return waitMethod ?? throw new InvalidOperationException(
            "Expected private method WaitForConnectionRefreshCompletion was not found.");
    }
}

[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ConnectionRefreshWaitPollingTests
{
    [Fact]
    public void RefreshWait_WhenCancellationRequested_InvokesCancelActionAndThrows()
    {
        using var cancelled = new CancellationTokenSource();
        var cancellationThread = new Thread(() =>
        {
            Thread.Sleep(50);
            cancelled.Cancel();
        });
        cancellationThread.Start();

        var cancelCalled = false;
        try
        {
            var error = Assert.Throws<TargetInvocationException>(() =>
                ConnectionRefreshWaitTests.GetRefreshWaitMethod().Invoke(null,
                [
                    (Func<bool>)(() => true),
                    (Action)(() => cancelCalled = true),
                    cancelled.Token
                ]));

            Assert.IsType<OperationCanceledException>(error.InnerException);
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
        var pollCount = 0;
        var cancelCalled = false;

        ConnectionRefreshWaitTests.GetRefreshWaitMethod().Invoke(null,
        [
            (Func<bool>)(() => Interlocked.Increment(ref pollCount) == 1),
            (Action)(() => cancelCalled = true),
            CancellationToken.None
        ]);

        Assert.True(pollCount >= 2);
        Assert.False(cancelCalled);
    }

}

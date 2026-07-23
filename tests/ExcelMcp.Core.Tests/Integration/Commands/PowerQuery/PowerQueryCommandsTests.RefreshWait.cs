using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

public partial class PowerQueryCommandsTests
{
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
        var waitMethod = typeof(PowerQueryCommands).GetMethod(
            "WaitForRefreshCompletion",
            BindingFlags.NonPublic | BindingFlags.Static);

        return waitMethod ?? throw new InvalidOperationException(
            "Expected private method WaitForRefreshCompletion was not found.");
    }
}

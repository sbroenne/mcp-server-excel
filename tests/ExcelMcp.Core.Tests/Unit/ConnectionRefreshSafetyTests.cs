using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "Connection")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ConnectionRefreshSafetyTests
{
    [Fact]
    public void RefreshPolicy_BackgroundSettingUnsupported_RefreshesAndPollsWithoutRestoring()
    {
        bool refreshed = false;
        bool polled = false;
        var settings = new List<bool>();
        InvokeRefreshPolicy(
            () => refreshed = true,
            value =>
            {
                settings.Add(value);
                throw Marshal.GetExceptionForHR(unchecked((int)0x800A03EC))!;
            },
            () =>
            {
                polled = true;
                return false;
            });
        Assert.True(refreshed);
        Assert.True(polled);
        Assert.Equal([false], settings);
    }

    [Fact]
    public void RefreshPolicy_RefreshAndRestoreFail_PreservesRefreshFailure()
    {
        var refreshError = new InvalidOperationException("Synthetic refresh failure");
        var settings = new List<bool>();
        var error = Assert.Throws<TargetInvocationException>(() => InvokeRefreshPolicy(
            () => throw refreshError,
            value =>
            {
                settings.Add(value);
                if (value)
                    throw Marshal.GetExceptionForHR(unchecked((int)0x800A03EC))!;
            },
            () => false));
        Assert.Same(refreshError, error.InnerException);
        Assert.Equal([false, true], settings);
    }

    [Fact]
    public void RefreshPolicy_BackgroundSettingUnsupported_StatusFailurePropagates()
    {
        var statusError = new InvalidOperationException("Synthetic status failure");
        var error = Assert.Throws<TargetInvocationException>(() => InvokeRefreshPolicy(
            () => { },
            _ => throw Marshal.GetExceptionForHR(unchecked((int)0x800A03EC))!,
            () => throw statusError));
        Assert.Same(statusError, error.InnerException);
    }

    private static void InvokeRefreshPolicy(Action refresh, Action<bool> setBackground, Func<bool> isRefreshing)
    {
        var method = typeof(ConnectionCommands).Assembly
            .GetType("Sbroenne.ExcelMcp.Core.Commands.ConnectionRefreshHelpers", throwOnError: true)!
            .GetMethod("Refresh", BindingFlags.NonPublic | BindingFlags.Static)!;
        method.Invoke(null,
        [
            refresh,
            (Func<bool>)(() => true),
            setBackground,
            isRefreshing,
            (Action)(() => { }),
            CancellationToken.None,
            (Action<Func<bool>, Action, CancellationToken>)((status, _, _) => _ = status())
        ]);
    }

    [Fact]
    public void QueryTableRefresh_CancelledResult_ThrowsMeaningfulFailure()
    {
        var method = GetRefreshResultMethod();
        var error = Assert.Throws<TargetInvocationException>(() => method.Invoke(null, [false]));
        var cause = Assert.IsType<InvalidOperationException>(error.InnerException);
        Assert.Contains("cancelled", cause.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void QueryTableRefresh_CompletedResult_DoesNotThrow()
    {
        GetRefreshResultMethod().Invoke(null, [true]);
    }

    [Theory]
    [InlineData(typeof(ConnectionCommands), "WaitForConnectionRefreshCompletion")]
    [InlineData(typeof(PowerQueryCommands), "WaitForRefreshCompletion")]
    public void RefreshWait_CancelledBeforeFirstPoll_DoesNotReportCompletion(Type commands, string methodName)
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        bool cancelled = false;
        var method = commands.GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static)!;
        var error = Assert.Throws<TargetInvocationException>(() => method.Invoke(null,
        [
            (Func<bool>)(() => false),
            (Action)(() => cancelled = true),
            cts.Token
        ]));
        Assert.IsType<OperationCanceledException>(error.InnerException);
        Assert.True(cancelled);
    }

    [Theory]
    [InlineData(typeof(ConnectionCommands), "WaitForConnectionRefreshCompletion")]
    [InlineData(typeof(PowerQueryCommands), "WaitForRefreshCompletion")]
    public void RefreshWait_StatusFailure_IsNotTreatedAsCompletion(Type commands, string methodName)
    {
        var statusError = new InvalidOperationException("Synthetic provider status failure");
        var method = commands.GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static)!;
        var error = Assert.Throws<TargetInvocationException>(() => method.Invoke(null,
        [
            (Func<bool>)(() => throw statusError),
            (Action)(() => { }),
            CancellationToken.None
        ]));
        Assert.Same(statusError, error.InnerException);
    }

    private static MethodInfo GetRefreshResultMethod() =>
        typeof(ConnectionCommands).Assembly
            .GetType("Sbroenne.ExcelMcp.Core.Commands.ConnectionRefreshHelpers", throwOnError: true)!
            .GetMethod("EnsureQueryTableRefreshSucceeded", BindingFlags.NonPublic | BindingFlags.Static)!;
}

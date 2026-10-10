using System.ComponentModel;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Collection("DpiAwareness")]
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Screenshot")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class DpiAwarenessTests
{
    [Fact]
    public void Execute_UsesPhysicalCoordinatesAndRestoresCallingThread()
    {
        IntPtr previous = NativeMethods.SetThreadDpiAwarenessContext(new IntPtr(-1));
        Assert.NotEqual(IntPtr.Zero, previous);
        try
        {
            int result = DpiAwareness.Execute(() =>
            {
                Assert.True(NativeMethods.AreDpiAwarenessContextsEqual(
                    new IntPtr(-4), NativeMethods.GetThreadDpiAwarenessContext()));
                return 42;
            });

            Assert.Equal(42, result);
            Assert.True(NativeMethods.AreDpiAwarenessContextsEqual(
                new IntPtr(-1), NativeMethods.GetThreadDpiAwarenessContext()));
        }
        finally
        {
            Assert.NotEqual(IntPtr.Zero, NativeMethods.SetThreadDpiAwarenessContext(previous));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_WhenOperationFails_PreservesFailureAndRestoresCallingThread(bool cancellation)
    {
        IntPtr previous = NativeMethods.SetThreadDpiAwarenessContext(new IntPtr(-1));
        Assert.NotEqual(IntPtr.Zero, previous);
        try
        {
            Exception expected = cancellation
                ? new OperationCanceledException("Synthetic capture cancellation.")
                : new InvalidOperationException("Synthetic capture failure.");
            var actual = Record.Exception(() => DpiAwareness.Execute<int>(() =>
            {
                Assert.True(NativeMethods.AreDpiAwarenessContextsEqual(
                    new IntPtr(-4), NativeMethods.GetThreadDpiAwarenessContext()));
                throw expected;
            }));

            Assert.Same(expected, actual);
            Assert.True(NativeMethods.AreDpiAwarenessContextsEqual(
                new IntPtr(-1), NativeMethods.GetThreadDpiAwarenessContext()));
        }
        finally
        {
            Assert.NotEqual(IntPtr.Zero, NativeMethods.SetThreadDpiAwarenessContext(previous));
        }
    }

    [Fact]
    public void Execute_NestedOperation_RestoresEachCallingContext()
    {
        IntPtr previous = NativeMethods.GetThreadDpiAwarenessContext();
        int result = DpiAwareness.Execute(() =>
        {
            int nested = DpiAwareness.Execute(() => 42);
            Assert.True(NativeMethods.AreDpiAwarenessContextsEqual(
                new IntPtr(-4), NativeMethods.GetThreadDpiAwarenessContext()));
            return nested;
        });

        Assert.Equal(42, result);
        Assert.True(NativeMethods.AreDpiAwarenessContextsEqual(
            previous, NativeMethods.GetThreadDpiAwarenessContext()));
    }

    [Fact]
    public void Execute_WhenCancellationAndRestorationFail_PreservesCancellationAndBothDiagnostics()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var primary = new OperationCanceledException("Synthetic capture cancellation.", cancellation.Token);
        var contexts = new List<IntPtr>();

        var actual = Assert.Throws<OperationCanceledException>(() => DpiAwareness.Execute<int>(
            () => throw primary,
            context =>
            {
                contexts.Add(context);
                if (contexts.Count == 1) { return new IntPtr(-1); }
                Marshal.SetLastPInvokeError(5);
                return IntPtr.Zero;
            }));

        Assert.Equal([new IntPtr(-4), new IntPtr(-1)], contexts);
        Assert.Equal(cancellation.Token, actual.CancellationToken);
        Assert.Equal("Cancelled", OperationFailureClassifier.Classify(actual));
        var failures = Assert.IsType<AggregateException>(actual.InnerException);
        Assert.Equal(2, failures.InnerExceptions.Count);
        Assert.Same(primary, failures.InnerExceptions[0]);
        var restoration = Assert.IsType<Win32Exception>(failures.InnerExceptions[1]);
        Assert.Equal(5, restoration.NativeErrorCode);
        Assert.Contains("restore the thread DPI awareness context", restoration.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_WhenRestorationFails_PreservesOperationAndNativeDiagnostics(bool operationFails)
    {
        var primary = new InvalidOperationException("Synthetic capture failure.");
        var contexts = new List<IntPtr>();

        var actual = Record.Exception(() => DpiAwareness.Execute(
            () => operationFails ? throw primary : 42,
            context =>
            {
                contexts.Add(context);
                if (contexts.Count == 1) { return new IntPtr(-1); }
                Marshal.SetLastPInvokeError(5);
                return IntPtr.Zero;
            }));

        Assert.Equal([new IntPtr(-4), new IntPtr(-1)], contexts);
        Exception? restoration = actual;
        if (operationFails)
        {
            var failures = Assert.IsType<AggregateException>(actual);
            Assert.Equal(2, failures.InnerExceptions.Count);
            Assert.Same(primary, failures.InnerExceptions[0]);
            restoration = failures.InnerExceptions[1];
        }
        Assert.Equal(5, Assert.IsType<Win32Exception>(restoration).NativeErrorCode);
    }

    private static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetThreadDpiAwarenessContext();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool AreDpiAwarenessContextsEqual(IntPtr first, IntPtr second);
    }
}

[CollectionDefinition("DpiAwareness", DisableParallelization = true)]
public sealed class DpiAwarenessTestGroup;

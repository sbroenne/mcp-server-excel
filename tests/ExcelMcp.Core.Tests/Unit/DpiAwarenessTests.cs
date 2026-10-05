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

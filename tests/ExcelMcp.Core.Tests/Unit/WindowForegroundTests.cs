using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Commands.Window;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "Window")]
public sealed class WindowForegroundTests
{
    [Fact]
    public void Activation_WhenDelayed_WaitsForObservedForeground()
    {
        var target = new IntPtr(42);
        var requested = IntPtr.Zero;
        var observations = 0;

        var result = WindowCommands.TryActivateWindow(
            target,
            hwnd =>
            {
                requested = hwnd;
                return true;
            },
            () => ++observations >= 3 ? target : IntPtr.Zero);

        Assert.True(result);
        Assert.Equal(target, requested);
        Assert.Equal(3, observations);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Activation_WhenNeverForeground_ReturnsFalse(bool requestAccepted)
    {
        var elapsed = Stopwatch.StartNew();

        var result = WindowCommands.TryActivateWindow(
            new IntPtr(42),
            _ => requestAccepted,
            () => new IntPtr(43));

        Assert.False(result);
        Assert.InRange(elapsed.Elapsed, TimeSpan.FromMilliseconds(900), TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void Activation_WhenAlreadyForeground_ReturnsTrueDespiteRefusedRequest()
    {
        var target = new IntPtr(42);

        var result = WindowCommands.TryActivateWindow(target, _ => false, () => target);

        Assert.True(result);
    }
}

// <copyright file="WindowCommandsTests.Arrange.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.ComponentModel;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Window;

/// <summary>
/// Tests for Arrange preset operations.
/// </summary>
public partial class WindowCommandsTests
{
    [Theory]
    [InlineData("left-half")]
    [InlineData("right-half")]
    [InlineData("top-half")]
    [InlineData("bottom-half")]
    [InlineData("center")]
    [InlineData("full-screen")]
    public void Arrange_ValidPresets_Succeed(string preset)
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Arrange(batch, preset);

        // Assert
        Assert.True(result.Success, $"Arrange '{preset}' failed: {result.ErrorMessage}");
        Assert.Equal("arrange", result.Action);
        Assert.Contains(preset, result.Message, StringComparison.OrdinalIgnoreCase);

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void Arrange_InvalidPreset_Throws()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert
        Assert.ThrowsAny<Exception>(() => _commands.Arrange(batch, "invalid-preset"));
    }

    [Theory]
    [InlineData("left-half")]
    [InlineData("right-half")]
    [InlineData("top-half")]
    [InlineData("bottom-half")]
    [InlineData("center")]
    public void Arrange_PresetUsesExcelMonitorWorkArea(string preset)
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        var workArea = GetExcelMonitorWorkArea(batch);

        // Act
        var result = _commands.Arrange(batch, preset);

        // Assert
        Assert.True(result.Success, result.ErrorMessage);

        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible);
        Assert.Equal("normal", info.WindowState);

        var expected = GetExpectedBounds(preset, workArea);
        AssertClose(expected.Left, info.Left);
        AssertClose(expected.Top, info.Top);
        AssertClose(expected.Width, info.Width);
        AssertClose(expected.Height, info.Height);

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void Arrange_FullScreen_MaximizesWindow()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Arrange(batch, "full-screen");

        // Assert
        Assert.True(result.Success);

        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible);
        Assert.Equal("maximized", info.WindowState);

        // Cleanup
        _commands.Hide(batch);
    }

    [Fact]
    public void Arrange_WhenHidden_MakesVisible()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        _commands.Hide(batch);

        // Act
        var result = _commands.Arrange(batch, "center");

        // Assert
        Assert.True(result.Success);

        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible, "Arrange should auto-show hidden Excel");

        // Cleanup
        _commands.Hide(batch);
    }

    private static Bounds GetExcelMonitorWorkArea(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var hwnd = new IntPtr(ctx.App.Hwnd);
            IntPtr previousDpiContext = SetThreadDpiAwarenessContext(PerMonitorAwareV2);
            Assert.NotEqual(IntPtr.Zero, previousDpiContext);
            try
            {
                var monitor = MonitorFromWindow(hwnd, MonitorDefaultToNearest);
                Assert.NotEqual(IntPtr.Zero, monitor);

                var info = new MonitorInfo
                {
                    Size = (uint)Marshal.SizeOf<MonitorInfo>()
                };
                if (!GetMonitorInfo(monitor, ref info))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                var dpi = GetDpiForWindow(hwnd);
                Assert.NotEqual(0u, dpi);
                var pointsPerPixel = 72d / dpi;
                return new Bounds(
                    info.WorkArea.Left * pointsPerPixel,
                    info.WorkArea.Top * pointsPerPixel,
                    (info.WorkArea.Right - info.WorkArea.Left) * pointsPerPixel,
                    (info.WorkArea.Bottom - info.WorkArea.Top) * pointsPerPixel);
            }
            finally
            {
                Assert.NotEqual(IntPtr.Zero, SetThreadDpiAwarenessContext(previousDpiContext));
            }
        });
    }

    private static Bounds GetExpectedBounds(string preset, Bounds workArea)
    {
        return preset switch
        {
            "left-half" => new(workArea.Left, workArea.Top, workArea.Width / 2, workArea.Height),
            "right-half" => new(
                workArea.Left + (workArea.Width / 2),
                workArea.Top,
                workArea.Width / 2,
                workArea.Height),
            "top-half" => new(workArea.Left, workArea.Top, workArea.Width, workArea.Height / 2),
            "bottom-half" => new(
                workArea.Left,
                workArea.Top + (workArea.Height / 2),
                workArea.Width,
                workArea.Height / 2),
            "center" => new(
                workArea.Left + (workArea.Width * 0.2),
                workArea.Top + (workArea.Height * 0.2),
                workArea.Width * 0.6,
                workArea.Height * 0.6),
            _ => throw new ArgumentOutOfRangeException(nameof(preset), preset, null)
        };
    }

    private static void AssertClose(double expected, double actual)
    {
        Assert.InRange(actual, expected - 8, expected + 8);
    }

    private const uint MonitorDefaultToNearest = 2;
    private static readonly IntPtr PerMonitorAwareV2 = new(-4);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

    [StructLayout(LayoutKind.Sequential)]
    private struct MonitorInfo
    {
        public uint Size;
        public NativeRect Monitor;
        public NativeRect WorkArea;
        public uint Flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private sealed record Bounds(double Left, double Top, double Width, double Height);
}

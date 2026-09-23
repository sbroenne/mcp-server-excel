using System.ComponentModel;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

/// <summary>
/// Resolves the monitor work area for a native window in Excel's point coordinate system.
/// </summary>
internal static class WindowWorkArea
{
    private const uint MonitorDefaultToNearest = 2;
    private const double PointsPerInch = 72;
    private static readonly IntPtr PerMonitorAwareV2 = new(-4);

    public static WindowBounds GetBoundsInPoints(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            throw new ArgumentException("A valid Excel window handle is required.", nameof(hwnd));
        }

        IntPtr previousDpiContext = SetThreadDpiAwarenessContext(PerMonitorAwareV2);
        if (previousDpiContext == IntPtr.Zero)
        {
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "Could not enable per-monitor DPI awareness while resolving the Excel monitor.");
        }

        WindowBounds bounds = default;
        Exception? failure = null;
        try
        {
            bounds = ResolveBounds(hwnd);
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            if (SetThreadDpiAwarenessContext(previousDpiContext) == IntPtr.Zero && failure is null)
            {
                failure = new Win32Exception(
                    Marshal.GetLastWin32Error(),
                    "Could not restore the thread DPI awareness context.");
            }
        }

        if (failure is not null)
        {
            ExceptionDispatchInfo.Capture(failure).Throw();
        }

        return bounds;
    }

    private static WindowBounds ResolveBounds(IntPtr hwnd)
    {
        IntPtr monitor = MonitorFromWindow(hwnd, MonitorDefaultToNearest);
        if (monitor == IntPtr.Zero)
        {
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "Could not identify the monitor containing the Excel window.");
        }

        var monitorInfo = new MonitorInfo
        {
            Size = (uint)Marshal.SizeOf<MonitorInfo>()
        };
        if (!GetMonitorInfo(monitor, ref monitorInfo))
        {
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "Could not read the work area of the monitor containing the Excel window.");
        }

        uint dpi = GetDpiForWindow(hwnd);
        if (dpi == 0)
        {
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "Could not read the DPI of the monitor containing the Excel window.");
        }

        double pointsPerPixel = PointsPerInch / dpi;
        return new WindowBounds(
            monitorInfo.WorkArea.Left * pointsPerPixel,
            monitorInfo.WorkArea.Top * pointsPerPixel,
            (monitorInfo.WorkArea.Right - monitorInfo.WorkArea.Left) * pointsPerPixel,
            (monitorInfo.WorkArea.Bottom - monitorInfo.WorkArea.Top) * pointsPerPixel);
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

    [DllImport("user32.dll", SetLastError = true)]
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
}

internal readonly record struct WindowBounds(double Left, double Top, double Width, double Height);

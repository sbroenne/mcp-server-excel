using System.ComponentModel;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

/// <summary>
/// Resolves the monitor work area for a native window in Excel's point coordinate system.
/// </summary>
internal static class WindowWorkArea
{
    private const uint MonitorDefaultToNearest = 2;
    private const double PointsPerInch = 72;

    public static WindowBounds GetBoundsInPoints(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            throw new ArgumentException("A valid Excel window handle is required.", nameof(hwnd));
        }

        return DpiAwareness.Execute(() => ResolveBounds(hwnd));
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

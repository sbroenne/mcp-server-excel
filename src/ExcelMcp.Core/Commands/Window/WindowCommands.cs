using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

/// <summary>
/// Implementation of window management commands using Excel COM and Win32 P/Invoke.
/// </summary>
public partial class WindowCommands : IWindowCommands
{
    // Win32 P/Invoke for window management
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    // Excel WindowState constants (XlWindowState)
    private const int XlMaximized = -4137;  // xlMaximized
    private const int XlMinimized = -4140;  // xlMinimized
    private const int XlNormal = -4143;     // xlNormal
    private const int SwRestore = 9;
    private const double WindowPlacementTolerance = 8;

    internal static Func<IntPtr, bool>? ForegroundActivationForTests { get; set; }

    /// <summary>
    /// Makes the Excel window visible and brings it to the foreground.
    /// </summary>
    public WindowVisibilityResult Show(IExcelBatch batch)
    {
        var result = GetWindowControl(batch).SetVisibility(true);
        return new WindowVisibilityResult
        {
            Success = true,
            Action = "show",
            IsVisible = result.IsVisible,
            IsForeground = result.IsForeground,
            Message = result.IsForeground ? "Excel is visible and in the foreground." :
                "Excel is visible, but Windows did not grant foreground focus. Select its window or taskbar entry."
        };
    }

    /// <summary>
    /// Hides the Excel window.
    /// </summary>
    public WindowVisibilityResult Hide(IExcelBatch batch)
    {
        var result = GetWindowControl(batch).SetVisibility(false);
        return new WindowVisibilityResult
        {
            Success = true,
            Action = "hide",
            IsVisible = result.IsVisible,
            IsForeground = result.IsForeground,
            Message = "Excel window is now hidden."
        };
    }

    private static IExcelBatchWindowControl GetWindowControl(IExcelBatch batch) =>
        batch as IExcelBatchWindowControl ??
        throw new NotSupportedException("This Excel session does not support native window visibility control.");

    /// <summary>
    /// Brings the Excel window to the foreground without changing visibility.
    /// </summary>
    public OperationResult BringToFront(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (!(bool)ctx.App.Visible)
            {
                return new OperationResult
                {
                    Success = true,
                    Action = "bring-to-front",
                    Message = "Excel is hidden. Use 'show' first to make it visible before bringing to front."
                };
            }

            if (!TryBringWindowToFront(ctx.App))
            {
                throw new InvalidOperationException(
                    "Windows could not bring the Excel window to the foreground. " +
                    "Select Excel on the Windows taskbar, or press Alt+Tab to switch to Excel.");
            }

            return new OperationResult
            {
                Success = true,
                Action = "bring-to-front",
                Message = "Excel window is now in the foreground"
            };
        });
    }

    /// <summary>
    /// Gets current window information.
    /// </summary>
    public WindowInfoResult GetInfo(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            bool isVisible = (bool)ctx.App.Visible;

            // Read window properties - these may throw if Excel is minimized or hidden
            double left = 0, top = 0, width = 0, height = 0;
            string windowState = "normal";

            if (isVisible)
            {
                int state = (int)ctx.App.WindowState;
                windowState = state switch
                {
                    XlMaximized => "maximized",
                    XlMinimized => "minimized",
                    _ => "normal"
                };

                left = Convert.ToDouble(ctx.App.Left);
                top = Convert.ToDouble(ctx.App.Top);
                width = Convert.ToDouble(ctx.App.Width);
                height = Convert.ToDouble(ctx.App.Height);
            }

            // Check if this is the foreground window
            int hwnd = ctx.App.Hwnd;
            IntPtr foreground = GetForegroundWindow();
            bool isForeground = isVisible && foreground == new IntPtr(hwnd);

            return new WindowInfoResult
            {
                Success = true,
                Action = "get-info",
                IsVisible = isVisible,
                WindowState = windowState,
                Left = left,
                Top = top,
                Width = width,
                Height = height,
                IsForeground = isForeground,
                Message = isVisible
                    ? $"Excel is visible ({windowState}), position: ({left},{top}), size: {width}x{height}"
                    : "Excel is hidden"
            };
        });
    }

    /// <summary>
    /// Sets the window state (normal, minimized, maximized).
    /// </summary>
    public OperationResult SetState(IExcelBatch batch, string windowState)
    {
        return batch.Execute((ctx, ct) =>
        {
            int xlState = ParseWindowState(windowState);

            // Ensure Excel is visible before changing state
            if (!(bool)ctx.App.Visible)
            {
                ctx.App.Visible = true;
            }

            ctx.App.WindowState = (Excel.XlWindowState)xlState;

            return new OperationResult
            {
                Success = true,
                Action = "set-state",
                Message = $"Excel window state set to {windowState}"
            };
        });
    }

    /// <summary>
    /// Sets the window position and size.
    /// </summary>
    public OperationResult SetPosition(IExcelBatch batch, double? left = null, double? top = null, double? width = null, double? height = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Ensure Excel is visible
            if (!(bool)ctx.App.Visible)
            {
                ctx.App.Visible = true;
            }

            // Set to normal state so position/size can be changed
            if ((int)ctx.App.WindowState != XlNormal)
            {
                ctx.App.WindowState = (Excel.XlWindowState)XlNormal;
            }

            if (left.HasValue) ctx.App.Left = left.Value;
            if (top.HasValue) ctx.App.Top = top.Value;
            if (width.HasValue) ctx.App.Width = width.Value;
            if (height.HasValue) ctx.App.Height = height.Value;

            return new OperationResult
            {
                Success = true,
                Action = "set-position",
                Message = $"Excel window position updated"
            };
        });
    }

    /// <summary>
    /// Arranges the Excel window using a named preset position.
    /// </summary>
    public WindowArrangeResult Arrange(IExcelBatch batch, string preset)
    {
        var normalizedPreset = preset.ToLowerInvariant();
        if (normalizedPreset is not ("left-half" or "right-half" or "top-half" or
            "bottom-half" or "center" or "full-screen"))
        {
            throw new ArgumentException(
                $"Unknown arrange preset: '{preset}'. " +
                "Valid presets: left-half, right-half, top-half, bottom-half, center, full-screen");
        }

        return batch.Execute((ctx, ct) =>
        {
            // Ensure Excel is visible
            if (!(bool)ctx.App.Visible)
            {
                ctx.App.Visible = true;
            }

            IntPtr hwnd = new(ctx.App.Hwnd);
            WindowBounds workArea = WindowWorkArea.GetBoundsInPoints(hwnd);

            // Set to normal state so position/size can be changed
            if (normalizedPreset != "full-screen")
            {
                if ((int)ctx.App.WindowState != XlNormal)
                {
                    ctx.App.WindowState = (Excel.XlWindowState)XlNormal;
                }
            }

            WindowBounds? expectedBounds = null;
            switch (normalizedPreset)
            {
                case "left-half":
                    expectedBounds = new WindowBounds(
                        workArea.Left,
                        workArea.Top,
                        workArea.Width / 2,
                        workArea.Height);
                    SetWindowBounds(ctx.App, expectedBounds.Value);
                    break;

                case "right-half":
                    expectedBounds = new WindowBounds(
                        workArea.Left + (workArea.Width / 2),
                        workArea.Top,
                        workArea.Width / 2,
                        workArea.Height);
                    SetWindowBounds(ctx.App, expectedBounds.Value);
                    break;

                case "top-half":
                    expectedBounds = new WindowBounds(
                        workArea.Left,
                        workArea.Top,
                        workArea.Width,
                        workArea.Height / 2);
                    SetWindowBounds(ctx.App, expectedBounds.Value);
                    break;

                case "bottom-half":
                    expectedBounds = new WindowBounds(
                        workArea.Left,
                        workArea.Top + (workArea.Height / 2),
                        workArea.Width,
                        workArea.Height / 2);
                    SetWindowBounds(ctx.App, expectedBounds.Value);
                    break;

                case "center":
                    expectedBounds = new WindowBounds(
                        workArea.Left + (workArea.Width * 0.2),
                        workArea.Top + (workArea.Height * 0.2),
                        workArea.Width * 0.6,
                        workArea.Height * 0.6);
                    SetWindowBounds(ctx.App, expectedBounds.Value);
                    break;

                case "full-screen":
                    ctx.App.WindowState = (Excel.XlWindowState)XlMaximized;
                    break;

            }

            VerifyWindowPlacement(ctx.App, normalizedPreset, expectedBounds);
            var isForeground = TryBringWindowToFront(ctx.App);

            return new WindowArrangeResult
            {
                Success = true,
                Action = "arrange",
                IsForeground = isForeground,
                Message = isForeground
                    ? $"Excel window arranged to '{preset}' and brought to the foreground."
                    : $"Excel window arranged to '{preset}'. Windows did not grant foreground focus. Select its window or taskbar entry."
            };
        });
    }

    private static void VerifyWindowPlacement(
        Excel.Application app,
        string preset,
        WindowBounds? expectedBounds)
    {
        if (expectedBounds is null)
        {
            var actualState = (int)app.WindowState;
            if (actualState != XlMaximized)
            {
                throw new InvalidOperationException(
                    $"Excel did not apply the requested '{preset}' placement. " +
                    $"Expected a maximized window, actual state '{actualState}'.");
            }
            return;
        }

        var actualWindowState = (int)app.WindowState;
        var actual = new WindowBounds(
            Convert.ToDouble(app.Left),
            Convert.ToDouble(app.Top),
            Convert.ToDouble(app.Width),
            Convert.ToDouble(app.Height));
        var expected = expectedBounds.Value;
        if (actualWindowState != XlNormal ||
            !IsWithinTolerance(expected.Left, actual.Left) ||
            !IsWithinTolerance(expected.Top, actual.Top) ||
            !IsWithinTolerance(expected.Width, actual.Width) ||
            !IsWithinTolerance(expected.Height, actual.Height))
        {
            throw new InvalidOperationException(
                $"Excel did not apply the requested '{preset}' placement. " +
                $"Expected bounds ({expected.Left}, {expected.Top}, {expected.Width}, {expected.Height}); " +
                $"actual bounds ({actual.Left}, {actual.Top}, {actual.Width}, {actual.Height}) " +
                $"and state '{actualWindowState}'.");
        }
    }

    private static bool IsWithinTolerance(double expected, double actual) =>
        Math.Abs(expected - actual) <= WindowPlacementTolerance;

    private static void SetWindowBounds(Excel.Application app, WindowBounds bounds)
    {
        app.Left = bounds.Left;
        app.Top = bounds.Top;
        app.Width = bounds.Width;
        app.Height = bounds.Height;
    }

    /// <summary>
    /// Parses a window state string to the Excel COM constant.
    /// </summary>
    private static int ParseWindowState(string windowState)
    {
        return windowState.ToLowerInvariant() switch
        {
            "normal" => XlNormal,
            "minimized" => XlMinimized,
            "maximized" => XlMaximized,
            _ => throw new ArgumentException(
                $"Unknown window state: '{windowState}'. Valid states: normal, minimized, maximized")
        };
    }

    /// <summary>
    /// Sets the Excel status bar text.
    /// </summary>
    public OperationResult SetStatusBar(IExcelBatch batch, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            ctx.App.StatusBar = text;

            return new OperationResult
            {
                Success = true,
                Action = "set-status-bar",
                Message = $"Status bar set to: {text}"
            };
        });
    }

    /// <summary>
    /// Clears the Excel status bar, restoring the default text.
    /// </summary>
    public OperationResult ClearStatusBar(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            // A Boolean is displayed as "FALSE" by some Excel COM versions.
            ctx.App.StatusBar = Type.Missing;

            return new OperationResult
            {
                Success = true,
                Action = "clear-status-bar",
                Message = "Status bar restored to default"
            };
        });
    }

    /// <summary>
    /// Brings the Excel window to the foreground using Win32 SetForegroundWindow.
    /// </summary>
    private static bool TryBringWindowToFront(Excel.Application app)
    {
        var hwnd = new IntPtr(app.Hwnd);
        if (hwnd == IntPtr.Zero)
        {
            throw new InvalidOperationException("Excel has no window handle to bring to the foreground.");
        }
        if (app.WindowState == Excel.XlWindowState.xlMinimized)
        {
            // SW_RESTORE retains the pre-minimize normal or maximized state.
            ShowWindow(hwnd, SwRestore);
        }
        if (ForegroundActivationForTests is { } activationForTests)
            return activationForTests(hwnd);

        return TryActivateWindow(hwnd, SetForegroundWindow, GetForegroundWindow);
    }

    internal static bool TryActivateWindow(
        IntPtr hwnd,
        Func<IntPtr, bool> requestForeground,
        Func<IntPtr> getForegroundWindow)
    {
        _ = requestForeground(hwnd);
        // Cross-process activation can finish after SetForegroundWindow returns.
        return SpinWait.SpinUntil(() => getForegroundWindow() == hwnd, TimeSpan.FromSeconds(1));
    }
}

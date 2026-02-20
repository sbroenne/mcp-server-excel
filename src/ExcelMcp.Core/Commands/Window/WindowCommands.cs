using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Window;

/// <summary>
/// Implementation of window management commands using Excel COM and Win32 P/Invoke.
/// </summary>
public class WindowCommands : IWindowCommands
{
    // Win32 P/Invoke for window management
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    // Excel WindowState constants (XlWindowState)
    private const int XlMaximized = -4137;  // xlMaximized
    private const int XlMinimized = -4140;  // xlMinimized
    private const int XlNormal = -4143;     // xlNormal

    /// <summary>
    /// Makes the Excel window visible and brings it to the foreground.
    /// </summary>
    public OperationResult Show(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            ctx.App.Visible = true;

            BringWindowToFront(ctx.App);

            return new OperationResult
            {
                Success = true,
                Action = "show",
                Message = "Excel window is now visible and in the foreground"
            };
        });
    }

    /// <summary>
    /// Hides the Excel window.
    /// </summary>
    public OperationResult Hide(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            ctx.App.Visible = false;

            return new OperationResult
            {
                Success = true,
                Action = "hide",
                Message = "Excel window is now hidden"
            };
        });
    }

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

            BringWindowToFront(ctx.App);

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
    public OperationResult Arrange(IExcelBatch batch, string preset)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Ensure Excel is visible
            if (!(bool)ctx.App.Visible)
            {
                ctx.App.Visible = true;
            }

            // Get screen dimensions via Excel's UsableWidth/UsableHeight
            // These are in points and represent the available screen area
            double screenWidth = Convert.ToDouble(ctx.App.UsableWidth);
            double screenHeight = Convert.ToDouble(ctx.App.UsableHeight);

            // Set to normal state so position/size can be changed
            if (preset != "full-screen")
            {
                if ((int)ctx.App.WindowState != XlNormal)
                {
                    ctx.App.WindowState = (Excel.XlWindowState)XlNormal;
                }
            }

            switch (preset.ToLowerInvariant())
            {
                case "left-half":
                    ctx.App.Left = 0;
                    ctx.App.Top = 0;
                    ctx.App.Width = screenWidth / 2;
                    ctx.App.Height = screenHeight;
                    break;

                case "right-half":
                    ctx.App.Left = screenWidth / 2;
                    ctx.App.Top = 0;
                    ctx.App.Width = screenWidth / 2;
                    ctx.App.Height = screenHeight;
                    break;

                case "top-half":
                    ctx.App.Left = 0;
                    ctx.App.Top = 0;
                    ctx.App.Width = screenWidth;
                    ctx.App.Height = screenHeight / 2;
                    break;

                case "bottom-half":
                    ctx.App.Left = 0;
                    ctx.App.Top = screenHeight / 2;
                    ctx.App.Width = screenWidth;
                    ctx.App.Height = screenHeight / 2;
                    break;

                case "center":
                    double centerWidth = screenWidth * 0.6;
                    double centerHeight = screenHeight * 0.6;
                    ctx.App.Left = (screenWidth - centerWidth) / 2;
                    ctx.App.Top = (screenHeight - centerHeight) / 2;
                    ctx.App.Width = centerWidth;
                    ctx.App.Height = centerHeight;
                    break;

                case "full-screen":
                    ctx.App.WindowState = (Excel.XlWindowState)XlMaximized;
                    break;

                default:
                    throw new ArgumentException(
                        $"Unknown arrange preset: '{preset}'. " +
                        "Valid presets: left-half, right-half, top-half, bottom-half, center, full-screen");
            }

            BringWindowToFront(ctx.App);

            return new OperationResult
            {
                Success = true,
                Action = "arrange",
                Message = $"Excel window arranged to '{preset}'"
            };
        });
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
            ctx.App.StatusBar = false; // false restores default "Ready" text

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
    private static void BringWindowToFront(dynamic app)
    {
        int hwnd = app.Hwnd;
        if (hwnd != 0)
        {
            SetForegroundWindow(new IntPtr(hwnd));
        }
    }
}

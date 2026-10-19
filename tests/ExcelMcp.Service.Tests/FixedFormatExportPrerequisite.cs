using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Excel lays out PDF/XPS pages through the Windows default printer. When that
/// printer's driver cannot be opened, ExportAsFixedFormat shows a modal
/// "Printer Setup" dialog that no test can answer, so the test stalls.
/// </summary>
internal static class FixedFormatExportPrerequisite
{
    private const int ErrorInsufficientBuffer = 122;
    private const int ErrorFileNotFound = 2;

    internal static void RequireUsableDefaultPrinter()
    {
        var printer = GetDefaultPrinterName();
        if (printer is null)
        {
            return;
        }

        var context = CreateIC(null, printer, null, IntPtr.Zero);
        if (context == IntPtr.Zero)
        {
            var error = Marshal.GetLastPInvokeError();
            throw new InvalidOperationException(
                $"PDF/XPS export test was not started: Windows cannot open the default printer " +
                $"\"{printer}\" (Win32 error {error}). Excel would show its Printer Setup dialog " +
                "during export and the test would hang. Repair or reinstall that printer, or choose " +
                "a working default printer, then run the test again.");
        }

        DeleteDC(context);
    }

    private static string? GetDefaultPrinterName()
    {
        var size = 0;
        if (GetDefaultPrinter(null, ref size))
        {
            return null;
        }

        var error = Marshal.GetLastPInvokeError();
        if (error == ErrorFileNotFound)
        {
            return null;
        }

        if (error != ErrorInsufficientBuffer || size <= 0)
        {
            throw new InvalidOperationException(
                $"Could not read the Windows default printer (Win32 error {error}).");
        }

        var buffer = new char[size];
        if (!GetDefaultPrinter(buffer, ref size))
        {
            throw new InvalidOperationException(
                $"Could not read the Windows default printer (Win32 error {Marshal.GetLastPInvokeError()}).");
        }

        return new string(buffer, 0, Math.Max(0, size - 1));
    }

    [DllImport("winspool.drv", EntryPoint = "GetDefaultPrinterW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetDefaultPrinter([Out] char[]? buffer, ref int size);

    [DllImport("gdi32.dll", EntryPoint = "CreateICW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateIC(string? driver, string device, string? port, IntPtr devMode);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteDC(IntPtr context);
}

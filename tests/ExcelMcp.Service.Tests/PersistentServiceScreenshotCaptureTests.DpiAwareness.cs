using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class IsolatedServiceScreenshotTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Capture_DpiUnawareThread_PreservesEveryCellAndRestoresContext(bool captureSheet)
    {
        var sheetName = CreateScreenshotSheetName();
        PopulateHighContrastOffscreenSheet(sheetName, topLeftCell: "AB90");
        _fixture.RegisterSheetForCleanup(sheetName);
        var before = ReadScreenshotState(sheetName, "AB90:AE96");

        var reference = CaptureWithDpiContext(new IntPtr(-4));
        AssertImageLooksPopulated(reference);
        using var referenceStream = new MemoryStream(Convert.FromBase64String(reference.ImageBase64));
        using var referenceBitmap = new Bitmap(referenceStream);

        for (int attempt = 0; attempt < 2; attempt++)
        {
            var result = CaptureWithDpiContext(new IntPtr(-1));
            Assert.True(result.Success, result.ErrorMessage);
            Assert.True(string.IsNullOrEmpty(result.ErrorMessage), result.ErrorMessage);
            Assert.Equal(sheetName, result.SheetName);
            Assert.Equal("$AB$90:$AE$96", result.RangeAddress);
            Assert.Equal(reference.Width, result.Width);
            Assert.Equal(reference.Height, result.Height);
            Assert.Equal(reference.MimeType, result.MimeType);
            Assert.Equal(before, ReadScreenshotState(sheetName, "AB90:AE96"));

            using var stream = new MemoryStream(Convert.FromBase64String(result.ImageBase64));
            using var bitmap = new Bitmap(stream);
            for (int row = 0; row < 7; row++)
            {
                for (int column = 0; column < 4; column++)
                {
                    int x = (int)((column + 0.9) * bitmap.Width / 4);
                    int y = (int)((row + 0.9) * bitmap.Height / 7);
                    Assert.Equal(referenceBitmap.GetPixel(x, y).ToArgb(), bitmap.GetPixel(x, y).ToArgb());
                }
            }
        }

        ScreenshotResult CaptureWithDpiContext(IntPtr context)
        {
            var previous = _fixture.ExecuteRawVerification((_, _) =>
                ScreenshotDpiNativeMethods.SetThreadDpiAwarenessContext(context));
            Assert.NotEqual(IntPtr.Zero, previous);
            try
            {
                var result = captureSheet
                    ? _screenshotCommands.CaptureSheet(_fixture.BatchToken, sheetName, ScreenshotQuality.High)
                    : _screenshotCommands.CaptureRange(_fixture.BatchToken, sheetName, "AB90:AE96", ScreenshotQuality.High);
                Assert.True(_fixture.ExecuteRawVerification((_, _) =>
                    ScreenshotDpiNativeMethods.AreDpiAwarenessContextsEqual(
                        context, ScreenshotDpiNativeMethods.GetThreadDpiAwarenessContext())));
                return result;
            }
            finally
            {
                _fixture.ExecuteRawVerification((_, _) => Assert.NotEqual(IntPtr.Zero,
                    ScreenshotDpiNativeMethods.SetThreadDpiAwarenessContext(previous)));
            }
        }
    }

    private static class ScreenshotDpiNativeMethods
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

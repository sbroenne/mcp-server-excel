using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.Versioning;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Scales and encodes captured bitmaps.
/// Ported from the shared capture approach in sbroenne/mcp-windows.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class ScreenshotEncoder
{
    private const long JpegQuality = 85;

    /// <summary>
    /// Encoded screenshot bytes together with the final pixel dimensions.
    /// </summary>
    public readonly record struct EncodedImage(byte[] Data, int Width, int Height, string MimeType);

    /// <summary>
    /// Scales the bitmap according to the requested quality and encodes it.
    /// High produces PNG at full scale, Medium JPEG at 75%, Low JPEG at 50% — matching the
    /// documented size/fidelity trade-off.
    /// </summary>
    public static EncodedImage Encode(Bitmap source, ScreenshotQuality quality)
    {
        ArgumentNullException.ThrowIfNull(source);

        double scale = quality switch
        {
            ScreenshotQuality.Low => 0.5,
            ScreenshotQuality.Medium => 0.75,
            _ => 1.0
        };

        Bitmap? scaled = null;

        try
        {
            Bitmap toEncode = source;

            if (scale < 1.0)
            {
                int width = Math.Max(1, (int)Math.Round(source.Width * scale));
                int height = Math.Max(1, (int)Math.Round(source.Height * scale));

                scaled = new Bitmap(width, height, PixelFormat.Format32bppArgb);

                using (var graphics = Graphics.FromImage(scaled))
                {
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.DrawImage(source, 0, 0, width, height);
                }

                toEncode = scaled;
            }

            bool png = quality == ScreenshotQuality.High;
            byte[] data = png ? EncodeToPng(toEncode) : EncodeToJpeg(toEncode);

            return new EncodedImage(
                data,
                toEncode.Width,
                toEncode.Height,
                png ? "image/png" : "image/jpeg");
        }
        finally
        {
            scaled?.Dispose();
        }
    }

    private static byte[] EncodeToPng(Bitmap bitmap)
    {
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    private static byte[] EncodeToJpeg(Bitmap bitmap)
    {
        ImageCodecInfo? encoder = Array.Find(
            ImageCodecInfo.GetImageEncoders(),
            codec => codec.FormatID == ImageFormat.Jpeg.Guid);

        if (encoder == null)
        {
            throw new InvalidOperationException("The JPEG encoder is unavailable on this machine.");
        }

        using var parameters = new EncoderParameters(1);
        parameters.Param[0] = new EncoderParameter(Encoder.Quality, JpegQuality);

        using var stream = new MemoryStream();
        bitmap.Save(stream, encoder, parameters);
        return stream.ToArray();
    }
}

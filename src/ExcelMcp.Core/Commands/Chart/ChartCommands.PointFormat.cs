using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

public partial class ChartCommands
{
    /// <inheritdoc />
    public ChartPointFormatResult GetPointFormat(IExcelBatch batch, string chartName, int seriesIndex, int pointIndex) =>
        WithPoint(batch, chartName, seriesIndex, pointIndex, null);

    /// <inheritdoc />
    public ChartPointFormatResult SetPointFormat(IExcelBatch batch, string chartName, int seriesIndex, int pointIndex, ChartPointOptions pointOptions)
    {
        ArgumentNullException.ThrowIfNull(pointOptions);
        ValidateMaterialFormat(pointOptions.FillTransparency, pointOptions.LineWeight);
        if (pointOptions.FillColor != null) FormattingHelpers.ParseColor(pointOptions.FillColor);
        if (pointOptions.LineColor != null) FormattingHelpers.ParseColor(pointOptions.LineColor);
        if ((pointOptions.FillTransparency.HasValue && !double.IsFinite(pointOptions.FillTransparency.Value)) ||
            (pointOptions.LineWeight.HasValue && (!double.IsFinite(pointOptions.LineWeight.Value) || pointOptions.LineWeight.Value > float.MaxValue)))
            throw new ArgumentException("Point material values must be finite and within Excel's numeric range.", nameof(pointOptions));
        if (pointOptions.MarkerStyle.HasValue && (!Enum.IsDefined(pointOptions.MarkerStyle.Value) || pointOptions.MarkerStyle == MarkerStyle.Picture))
            throw new ArgumentException("Unknown marker style or unsupported Picture marker.", nameof(pointOptions));
        if (pointOptions.MarkerSize is < 2 or > 72)
            throw new ArgumentOutOfRangeException(nameof(pointOptions), "Marker size must be 2 to 72.");
        return WithPoint(batch, chartName, seriesIndex, pointIndex, pointOptions);
    }

    private ChartPointFormatResult WithPoint(IExcelBatch batch, string name, int seriesIndex, int pointIndex, ChartPointOptions? options)
    {
        ValidateSeriesIndex(seriesIndex);
        if (pointIndex < 1) throw new ArgumentOutOfRangeException(nameof(pointIndex), "Point indices start at one.");
        return WithNativeSeries(batch, name, seriesIndex, options != null, (_, series, _, _, _) =>
        {
            Excel.Points? points = null;
            Excel.Point? point = null;
            dynamic? format = null;
            try
            {
                points = (Excel.Points)series.Points();
                if (pointIndex > points.Count)
                    throw new ArgumentOutOfRangeException(nameof(pointIndex), $"Series has {points.Count} points.");
                var typeName = series.ChartType.ToString();
                var markers = typeName.Contains("Line", StringComparison.Ordinal) ||
                    typeName.Contains("Scatter", StringComparison.Ordinal) || typeName.Contains("Radar", StringComparison.Ordinal);
                if (!markers && options != null && (options.MarkerStyle.HasValue || options.MarkerSize.HasValue))
                    throw new ArgumentException("Marker settings require a line, scatter or radar series.", nameof(options));
                if (markers && options?.FillTransparency != null)
                    throw new NotSupportedException("Excel does not persist per-point marker fill transparency through this operation. Omit fillTransparency for line/scatter/radar points.");
                if (markers && options?.LineWeight != null)
                    throw new NotSupportedException("Per-point marker outline weight is not supported. Omit lineWeight for line/scatter/radar points.");
                point = points.Item(pointIndex);
                // PIA gap: chart material fill/line objects require unavailable Office.Core types.
                format = point.Format;
                if (options != null)
                {
                    if (options.MarkerStyle.HasValue) point.MarkerStyle = (Excel.XlMarkerStyle)options.MarkerStyle.Value;
                    if (options.MarkerSize.HasValue) point.MarkerSize = options.MarkerSize.Value;
                    if (markers)
                    {
                        if (options.FillColor != null) point.MarkerBackgroundColor = FormattingHelpers.ParseColor(options.FillColor);
                        if (options.LineColor != null) point.MarkerForegroundColor = FormattingHelpers.ParseColor(options.LineColor);
                    }
                    else
                    {
                        ApplyMaterialFormat(format, options.FillColor, options.FillTransparency, options.LineColor, options.LineWeight);
                    }
                }
                return ReadPointMaterial(batch, name, seriesIndex, pointIndex, point, format, markers);
            }
            finally
            {
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref point);
                ComUtilities.Release(ref points);
            }
        });
    }

    private static ChartPointFormatResult ReadPointMaterial(
        IExcelBatch batch, string name, int seriesIndex, int pointIndex, Excel.Point point, dynamic format, bool markers)
    {
        dynamic? fill = null;
        dynamic? fillColor = null;
        dynamic? line = null;
        dynamic? lineColor = null;
        try
        {
            // PIA gap: Office.Core fill/line/color getters need late binding.
            fill = format.Fill;
            fillColor = fill.ForeColor;
            line = format.Line;
            lineColor = line.ForeColor;
            var transparency = Convert.ToDouble(fill.Transparency);
            var transparencyAvailable = !markers && double.IsFinite(transparency) && transparency is >= 0 and <= 1;
            var fillRgb = markers ? point.MarkerBackgroundColor : Convert.ToInt32(fillColor.RGB);
            var lineRgb = markers ? point.MarkerForegroundColor : Convert.ToInt32(lineColor.RGB);
            var fillRgbAvailable = fillRgb is >= 0 and <= 0xFFFFFF;
            var lineRgbAvailable = lineRgb is >= 0 and <= 0xFFFFFF;
            return new ChartPointFormatResult
            {
                Success = true,
                FilePath = batch.WorkbookPath,
                ChartName = name,
                SeriesIndex = seriesIndex,
                PointIndex = pointIndex,
                FillColor = fillRgbAvailable ? FormattingHelpers.ColorToHex(fillRgb) : null,
                FillColorAvailable = fillRgbAvailable,
                FillColorReadError = fillRgbAvailable ? null : $"Excel returned an automatic/mixed color value ({fillRgb}); an explicit RGB fill cannot be inspected.",
                FillTransparency = transparencyAvailable ? transparency : null,
                FillTransparencyAvailable = transparencyAvailable,
                FillTransparencyReadError = transparencyAvailable ? null : $"Excel returned an invalid point transparency value ({transparency.ToString(System.Globalization.CultureInfo.InvariantCulture)}); the actual transparency cannot be inspected.",
                LineColor = lineRgbAvailable ? FormattingHelpers.ColorToHex(lineRgb) : null,
                LineColorAvailable = lineRgbAvailable,
                LineColorReadError = lineRgbAvailable ? null : $"Excel returned an automatic/mixed color value ({lineRgb}); an explicit RGB line cannot be inspected.",
                LineWeight = markers ? null : Convert.ToDouble(line.Weight),
                LineWeightAvailable = !markers,
                MarkersSupported = markers,
                MarkerStyle = markers ? (MarkerStyle)point.MarkerStyle : null,
                MarkerSize = markers ? point.MarkerSize : null
            };
        }
        finally
        {
            ComUtilities.Release(ref lineColor);
            ComUtilities.Release(ref line);
            ComUtilities.Release(ref fillColor);
            ComUtilities.Release(ref fill);
        }
    }

}

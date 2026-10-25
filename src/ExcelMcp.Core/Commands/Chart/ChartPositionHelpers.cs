using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Helpers for chart positioning: collision detection against data ranges and other charts.
/// </summary>
internal static class ChartPositionHelpers
{
    /// <summary>
    /// Detects collisions between a chart's proposed position and existing content on the worksheet.
    /// Checks against the used range (data) and all other chart shapes.
    /// </summary>
    /// <param name="worksheet">Target worksheet COM object</param>
    /// <param name="left">Proposed chart left position in points</param>
    /// <param name="top">Proposed chart top position in points</param>
    /// <param name="width">Proposed chart width in points</param>
    /// <param name="height">Proposed chart height in points</param>
    /// <param name="excludeChartName">Chart name to exclude from collision checks (for move/resize operations on an existing chart)</param>
    /// <returns>List of collision warning messages, empty if no collisions</returns>
    internal static List<string> DetectCollisions(
        Excel.Worksheet worksheet,
        double left,
        double top,
        double width,
        double height,
        string? excludeChartName = null)
    {
        var warnings = new List<string>();

        // Check collision with used range (data area)
        CheckUsedRangeCollision(worksheet, left, top, width, height, warnings);

        // Check collision with other charts
        CheckChartCollisions(worksheet, left, top, width, height, excludeChartName, warnings);

        return warnings;
    }

    /// <summary>
    /// Finds the first available position below or to the right of all existing content (used range + charts)
    /// on the worksheet for a chart with the given dimensions.
    /// </summary>
    /// <param name="worksheet">Target worksheet COM object</param>
    /// <param name="width">Desired chart width in points</param>
    /// <param name="height">Desired chart height in points</param>
    /// <param name="padding">Padding in points between content and chart (default: 10pt)</param>
    /// <returns>Tuple of (left, top) position in points for the chart</returns>
    internal static (double Left, double Top) FindAvailablePosition(
        Excel.Worksheet worksheet,
        double width = 400,
        double height = 300,
        double padding = 10.0)
    {
        double maxBottom = 0;
        double maxRight = 0;

        // Get used range boundary
        Excel.Range? usedRange = null;
        try
        {
            usedRange = worksheet.UsedRange;
            double urLeft = Convert.ToDouble((object)usedRange.Left, System.Globalization.CultureInfo.InvariantCulture);
            double urTop = Convert.ToDouble((object)usedRange.Top, System.Globalization.CultureInfo.InvariantCulture);
            double urWidth = Convert.ToDouble((object)usedRange.Width, System.Globalization.CultureInfo.InvariantCulture);
            double urHeight = Convert.ToDouble((object)usedRange.Height, System.Globalization.CultureInfo.InvariantCulture);

            maxBottom = Math.Max(maxBottom, urTop + urHeight);
            maxRight = Math.Max(maxRight, urLeft + urWidth);
        }
        finally
        {
            ComUtilities.Release(ref usedRange!);
        }

        // Get chart boundaries
        Excel.Shapes? shapes = null;
        try
        {
            shapes = worksheet.Shapes;
            int shapeCount = Convert.ToInt32(shapes.Count);

            for (int j = 1; j <= shapeCount; j++)
            {
                Excel.Shape? shape = null;
                try
                {
                    shape = shapes.Item(j);

                    // Only consider chart shapes (msoChart = 3)
                    if ((int)shape.Type != 3)
                    {
                        continue;
                    }

                    double sLeft = Convert.ToDouble(shape.Left);
                    double sTop = Convert.ToDouble(shape.Top);
                    double sWidth = Convert.ToDouble(shape.Width);
                    double sHeight = Convert.ToDouble(shape.Height);

                    maxBottom = Math.Max(maxBottom, sTop + sHeight);
                    maxRight = Math.Max(maxRight, sLeft + sWidth);
                }
                finally
                {
                    ComUtilities.Release(ref shape!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref shapes!);
        }

        // Strategy: place chart below all existing content, aligned to left edge of used range
        // If the chart would fit to the right of the used range (within reasonable horizontal space),
        // place it there instead for a side-by-side layout.
        // For now, always place below to avoid horizontal overflow.
        _ = width;  // Reserved for future side-by-side layout consideration
        _ = height; // Reserved for future vertical space check

        return (padding, maxBottom + padding);
    }

    /// <summary>
    /// Counts the number of chart shapes on a worksheet.
    /// </summary>
    internal static int CountCharts(Excel.Worksheet worksheet)
    {
        int count = 0;
        Excel.Shapes? shapes = null;
        try
        {
            shapes = worksheet.Shapes;
            int shapeCount = Convert.ToInt32(shapes.Count);

            for (int j = 1; j <= shapeCount; j++)
            {
                Excel.Shape? shape = null;
                try
                {
                    shape = shapes.Item(j);
                    if ((int)shape.Type == 3) // msoChart
                    {
                        count++;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shape!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref shapes!);
        }

        return count;
    }

    /// <summary>
    /// Formats chart positioning feedback for the result message.
    /// Includes overlap warnings and verification options appropriate to the available desktop.
    /// </summary>
    /// <param name="warnings">Collision warnings (empty if no overlaps detected)</param>
    /// <param name="chartCount">Number of charts on the worksheet (including the one just created/moved)</param>
    internal static string FormatCollisionWarnings(List<string> warnings, int chartCount = 1)
    {
        const string verification = "Inspect bounds with chart read; use screenshot capture-sheet for visual verification when an interactive desktop is available.";
        if (warnings.Count > 0)
        {
            return $"OVERLAP WARNING: {string.Join("; ", warnings)}. Use chart move or fit-to-range to reposition. {verification}";
        }

        if (chartCount >= 2)
        {
            return $"{chartCount} charts now on this sheet. {verification}";
        }

        return verification;
    }

    private static void CheckUsedRangeCollision(
        Excel.Worksheet worksheet,
        double left,
        double top,
        double width,
        double height,
        List<string> warnings)
    {
        Excel.Range? usedRange = null;
        try
        {
            usedRange = worksheet.UsedRange;

            double urLeft = Convert.ToDouble((object)usedRange.Left, System.Globalization.CultureInfo.InvariantCulture);
            double urTop = Convert.ToDouble((object)usedRange.Top, System.Globalization.CultureInfo.InvariantCulture);
            double urWidth = Convert.ToDouble((object)usedRange.Width, System.Globalization.CultureInfo.InvariantCulture);
            double urHeight = Convert.ToDouble((object)usedRange.Height, System.Globalization.CultureInfo.InvariantCulture);

            string? urAddress = usedRange.Address?.ToString();

            if (RectsOverlap(left, top, width, height, urLeft, urTop, urWidth, urHeight))
            {
                warnings.Add($"Chart overlaps data area {urAddress}");
            }
        }
        finally
        {
            ComUtilities.Release(ref usedRange!);
        }
    }

    private static void CheckChartCollisions(
        Excel.Worksheet worksheet,
        double left,
        double top,
        double width,
        double height,
        string? excludeChartName,
        List<string> warnings)
    {
        Excel.Shapes? shapes = null;
        try
        {
            shapes = worksheet.Shapes;
            int shapeCount = Convert.ToInt32(shapes.Count);

            for (int j = 1; j <= shapeCount; j++)
            {
                Excel.Shape? shape = null;
                try
                {
                    shape = shapes.Item(j);

                    // Only check chart shapes (msoChart = 3)
                    if ((int)shape.Type != 3)
                    {
                        continue;
                    }

                    string shapeName = shape.Name?.ToString() ?? string.Empty;

                    // Skip the chart being created/moved
                    if (excludeChartName != null &&
                        shapeName.Equals(excludeChartName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    double sLeft = Convert.ToDouble(shape.Left);
                    double sTop = Convert.ToDouble(shape.Top);
                    double sWidth = Convert.ToDouble(shape.Width);
                    double sHeight = Convert.ToDouble(shape.Height);

                    if (RectsOverlap(left, top, width, height, sLeft, sTop, sWidth, sHeight))
                    {
                        warnings.Add($"Chart overlaps existing chart '{shapeName}'");
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shape!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref shapes!);
        }
    }

    /// <summary>
    /// Checks if two rectangles overlap (in point coordinates).
    /// Returns false if they merely touch edges (share a boundary).
    /// </summary>
    private static bool RectsOverlap(
        double x1, double y1, double w1, double h1,
        double x2, double y2, double w2, double h2)
    {
        // No overlap if one is completely to the left, right, above, or below the other
        // Use strict < (not <=) so touching edges are NOT considered overlaps
        return x1 < x2 + w2 &&
               x1 + w1 > x2 &&
               y1 < y2 + h2 &&
               y1 + h1 > y2;
    }
}

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
        dynamic worksheet,
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
        dynamic worksheet,
        double width = 400,
        double height = 300,
        double padding = 10.0)
    {
        double maxBottom = 0;
        double maxRight = 0;

        // Get used range boundary
        dynamic? usedRange = null;
        try
        {
            usedRange = worksheet.UsedRange;
            double urLeft = Convert.ToDouble(usedRange.Left);
            double urTop = Convert.ToDouble(usedRange.Top);
            double urWidth = Convert.ToDouble(usedRange.Width);
            double urHeight = Convert.ToDouble(usedRange.Height);

            maxBottom = Math.Max(maxBottom, urTop + urHeight);
            maxRight = Math.Max(maxRight, urLeft + urWidth);
        }
        finally
        {
            ComUtilities.Release(ref usedRange!);
        }

        // Get chart boundaries
        dynamic? shapes = null;
        try
        {
            shapes = worksheet.Shapes;
            int shapeCount = Convert.ToInt32(shapes.Count);

            for (int j = 1; j <= shapeCount; j++)
            {
                dynamic? shape = null;
                try
                {
                    shape = shapes.Item(j);

                    // Only consider chart shapes (msoChart = 3)
                    if (Convert.ToInt32(shape.Type) != 3)
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
    internal static int CountCharts(dynamic worksheet)
    {
        int count = 0;
        dynamic? shapes = null;
        try
        {
            shapes = worksheet.Shapes;
            int shapeCount = Convert.ToInt32(shapes.Count);

            for (int j = 1; j <= shapeCount; j++)
            {
                dynamic? shape = null;
                try
                {
                    shape = shapes.Item(j);
                    if (Convert.ToInt32(shape.Type) == 3) // msoChart
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
    /// Always includes a screenshot verification reminder.
    /// When collisions are detected, includes overlap warnings with remediation guidance.
    /// When multiple charts exist on the sheet, uses stronger language to ensure screenshot verification.
    /// </summary>
    /// <param name="warnings">Collision warnings (empty if no overlaps detected)</param>
    /// <param name="chartCount">Number of charts on the worksheet (including the one just created/moved)</param>
    internal static string FormatCollisionWarnings(List<string> warnings, int chartCount = 1)
    {
        if (warnings.Count > 0)
        {
            return $"OVERLAP WARNING: {string.Join("; ", warnings)}. Use chart move or fit-to-range to reposition, then screenshot(capture-sheet) to verify layout.";
        }

        if (chartCount >= 2)
        {
            return $"IMPORTANT: {chartCount} charts now on this sheet. You MUST take a screenshot(capture-sheet) to verify no charts overlap each other or the data.";
        }

        return "IMPORTANT: You MUST take a screenshot(capture-sheet) to verify the chart does not overlap the data.";
    }

    private static void CheckUsedRangeCollision(
        dynamic worksheet,
        double left,
        double top,
        double width,
        double height,
        List<string> warnings)
    {
        dynamic? usedRange = null;
        try
        {
            usedRange = worksheet.UsedRange;

            double urLeft = Convert.ToDouble(usedRange.Left);
            double urTop = Convert.ToDouble(usedRange.Top);
            double urWidth = Convert.ToDouble(usedRange.Width);
            double urHeight = Convert.ToDouble(usedRange.Height);

            string? urAddress = null;
            try
            {
                urAddress = usedRange.Address?.ToString();
            }
            catch
            {
                urAddress = "(unknown)";
            }

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
        dynamic worksheet,
        double left,
        double top,
        double width,
        double height,
        string? excludeChartName,
        List<string> warnings)
    {
        dynamic? shapes = null;
        try
        {
            shapes = worksheet.Shapes;
            int shapeCount = Convert.ToInt32(shapes.Count);

            for (int j = 1; j <= shapeCount; j++)
            {
                dynamic? shape = null;
                try
                {
                    shape = shapes.Item(j);

                    // Only check chart shapes (msoChart = 3)
                    if (Convert.ToInt32(shape.Type) != 3)
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

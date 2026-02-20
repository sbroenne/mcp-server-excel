using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart appearance operations - type, title, axes, legend, style.
/// </summary>
public partial class ChartCommands
{
    /// <inheritdoc />
    public OperationResult SetChartType(IExcelBatch batch, string chartName, ChartType chartType)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            try
            {
                // Set chart type (works for both Regular and PivotCharts)
                findResult.Chart.ChartType = (int)chartType;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetTitle(IExcelBatch batch, string chartName, string title)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            try
            {
                // Set title (empty string hides title)
                if (string.IsNullOrEmpty(title))
                {
                    findResult.Chart.HasTitle = false;
                }
                else
                {
                    findResult.Chart.HasTitle = true;
                    findResult.Chart.ChartTitle.Text = title;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetAxisTitle(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        string title)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? axes = null;
            dynamic? targetAxis = null;

            try
            {
                axes = findResult.Chart.Axes;

                // Map axis type to Excel constants
                int axisType = axis switch
                {
                    ChartAxisType.Category => 1,    // xlCategory
                    ChartAxisType.Value => 2,       // xlValue
                    ChartAxisType.Primary => 1,     // Primary = Category
                    ChartAxisType.Secondary => 2,   // Secondary = Value
                    _ => 1
                };

                targetAxis = axes.Item(axisType);

                // Set axis title (empty string hides title)
                if (string.IsNullOrEmpty(title))
                {
                    targetAxis.HasTitle = false;
                }
                else
                {
                    targetAxis.HasTitle = true;
                    targetAxis.AxisTitle.Text = title;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                ComUtilities.Release(ref targetAxis!);
                ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public string GetAxisNumberFormat(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? axes = null;
            dynamic? targetAxis = null;
            dynamic? tickLabels = null;

            try
            {
                axes = findResult.Chart.Axes;

                // Map axis type to Excel constants
                int axisType = axis switch
                {
                    ChartAxisType.Category => 1,    // xlCategory
                    ChartAxisType.Value => 2,       // xlValue
                    ChartAxisType.Primary => 1,     // Primary = Category
                    ChartAxisType.Secondary => 2,   // Secondary = Value
                    _ => 1
                };

                targetAxis = axes.Item(axisType);
                tickLabels = targetAxis.TickLabels;

                // Get the number format for axis tick labels
                return tickLabels.NumberFormat?.ToString() ?? "General";
            }
            finally
            {
                ComUtilities.Release(ref tickLabels!);
                ComUtilities.Release(ref targetAxis!);
                ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetAxisNumberFormat(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        string numberFormat)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? axes = null;
            dynamic? targetAxis = null;
            dynamic? tickLabels = null;

            try
            {
                axes = findResult.Chart.Axes;

                // Map axis type to Excel constants
                int axisType = axis switch
                {
                    ChartAxisType.Category => 1,    // xlCategory
                    ChartAxisType.Value => 2,       // xlValue
                    ChartAxisType.Primary => 1,     // Primary = Category
                    ChartAxisType.Secondary => 2,   // Secondary = Value
                    _ => 1
                };

                targetAxis = axes.Item(axisType);
                tickLabels = targetAxis.TickLabels;

                // Set the number format for axis tick labels
                tickLabels.NumberFormat = numberFormat;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                ComUtilities.Release(ref tickLabels!);
                ComUtilities.Release(ref targetAxis!);
                ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ShowLegend(
        IExcelBatch batch,
        string chartName,
        bool visible,
        LegendPosition? legendPosition = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? legend = null;

            try
            {
                // Show/hide legend
                findResult.Chart.HasLegend = visible;

                // Set position if provided and legend is visible
                if (visible && legendPosition.HasValue)
                {
                    legend = findResult.Chart.Legend;
                    legend.Position = (int)legendPosition.Value;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                ComUtilities.Release(ref legend!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetStyle(IExcelBatch batch, string chartName, int styleId)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            try
            {
                // Validate range (Excel supports styles 1-48)
                if (styleId < 1 || styleId > 48)
                {
                    var hint = styleId == 0 ? " (was the 'style_id' parameter included?)" : "";
                    throw new ArgumentException($"Chart style ID must be between 1 and 48. Provided: {styleId}{hint}", nameof(styleId));
                }

                // Set chart style
                findResult.Chart.ChartStyle = styleId;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetPlacement(IExcelBatch batch, string chartName, int placement)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            try
            {
                // Validate placement value (xlMoveAndSize=1, xlMove=2, xlFreeFloating=3)
                if (placement < 1 || placement > 3)
                {
                    throw new ArgumentException(
                        $"Placement must be 1 (move and size with cells), 2 (move only), or 3 (free floating). Provided: {placement}",
                        nameof(placement));
                }

                // Set placement on the shape (ChartObject)
                findResult.Shape.Placement = placement;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    // === DATA LABELS ===

    /// <inheritdoc />
    public OperationResult SetDataLabels(
        IExcelBatch batch,
        string chartName,
        bool? showValue = null,
        bool? showPercentage = null,
        bool? showSeriesName = null,
        bool? showCategoryName = null,
        bool? showBubbleSize = null,
        string? separator = null,
        DataLabelPosition? labelPosition = null,
        int? seriesIndex = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? seriesCollection = null;
            dynamic? series = null;
            dynamic? dataLabels = null;

            try
            {
                seriesCollection = findResult.Chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                if (seriesCount == 0)
                {
                    throw new InvalidOperationException($"Chart '{chartName}' has no data series.");
                }

                // Treat 0 as "all series" (MCP clients may send 0 when parameter is omitted)
                if (seriesIndex == 0) seriesIndex = null;

                // Determine which series to configure
                int startIndex = seriesIndex ?? 1;
                int endIndex = seriesIndex ?? seriesCount;

                if (seriesIndex.HasValue && (seriesIndex.Value < 1 || seriesIndex.Value > seriesCount))
                {
                    throw new ArgumentException($"Series index {seriesIndex.Value} is out of range. Chart has {seriesCount} series (1-based).");
                }

                for (int i = startIndex; i <= endIndex; i++)
                {
                    series = seriesCollection.Item(i);

                    // First enable data labels if any property is being set
                    if (showValue == true || showPercentage == true || showSeriesName == true ||
                        showCategoryName == true || showBubbleSize == true)
                    {
                        series.HasDataLabels = true;
                    }

                    dataLabels = series.DataLabels;

                    // Apply each property if specified
                    if (showValue.HasValue)
                        dataLabels.ShowValue = showValue.Value;

                    if (showPercentage.HasValue)
                    {
                        try
                        {
                            dataLabels.ShowPercentage = showPercentage.Value;
                        }
                        catch (System.Runtime.InteropServices.COMException ex)
                            when (ex.HResult == unchecked((int)0x800A03EC))
                        {
                            throw new InvalidOperationException(
                                $"ShowPercentage is not supported for this chart type. " +
                                "Use show_percentage only with pie or doughnut chart types.", ex);
                        }
                    }

                    if (showSeriesName.HasValue)
                        dataLabels.ShowSeriesName = showSeriesName.Value;

                    if (showCategoryName.HasValue)
                        dataLabels.ShowCategoryName = showCategoryName.Value;

                    if (showBubbleSize.HasValue)
                        dataLabels.ShowBubbleSize = showBubbleSize.Value;

                    if (!string.IsNullOrEmpty(separator))
                        dataLabels.Separator = separator;

                    if (labelPosition.HasValue)
                        dataLabels.Position = (int)labelPosition.Value;

                    // Disable data labels entirely if all show properties are false
                    if (showValue == false && showPercentage == false && showSeriesName == false &&
                        showCategoryName == false && showBubbleSize == false)
                    {
                        series.HasDataLabels = false;
                    }

                    ComUtilities.Release(ref dataLabels!);
                    dataLabels = null;
                    ComUtilities.Release(ref series!);
                    series = null;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref dataLabels!);
                ComUtilities.Release(ref series!);
                ComUtilities.Release(ref seriesCollection!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    // === AXIS SCALE ===

    /// <inheritdoc />
    public AxisScaleResult GetAxisScale(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? axes = null;
            dynamic? targetAxis = null;

            try
            {
                axes = findResult.Chart.Axes;

                // Map axis type to Excel constants
                var (axisType, axisGroup) = MapAxisType(axis);
                targetAxis = axes.Item(axisType, axisGroup);

                var result = new AxisScaleResult
                {
                    Success = true,
                    ChartName = chartName,
                    AxisType = axis.ToString()
                };

                // Get scale properties with safe null handling
                result.MinimumScaleIsAuto = targetAxis.MinimumScaleIsAuto;
                result.MaximumScaleIsAuto = targetAxis.MaximumScaleIsAuto;
                result.MajorUnitIsAuto = targetAxis.MajorUnitIsAuto;
                result.MinorUnitIsAuto = targetAxis.MinorUnitIsAuto;

                if (!result.MinimumScaleIsAuto)
                    result.MinimumScale = targetAxis.MinimumScale;

                if (!result.MaximumScaleIsAuto)
                    result.MaximumScale = targetAxis.MaximumScale;

                if (!result.MajorUnitIsAuto)
                    result.MajorUnit = targetAxis.MajorUnit;

                if (!result.MinorUnitIsAuto)
                    result.MinorUnit = targetAxis.MinorUnit;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetAxis!);
                ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetAxisScale(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        double? minimumScale = null,
        double? maximumScale = null,
        double? majorUnit = null,
        double? minorUnit = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? axes = null;
            dynamic? targetAxis = null;

            try
            {
                axes = findResult.Chart.Axes;

                // Map axis type to Excel constants
                var (axisType, axisGroup) = MapAxisType(axis);
                targetAxis = axes.Item(axisType, axisGroup);

                // Set scale properties
                // If value is provided, use it; otherwise, set to auto
                if (minimumScale.HasValue)
                {
                    targetAxis.MinimumScaleIsAuto = false;
                    targetAxis.MinimumScale = minimumScale.Value;
                }

                if (maximumScale.HasValue)
                {
                    targetAxis.MaximumScaleIsAuto = false;
                    targetAxis.MaximumScale = maximumScale.Value;
                }

                if (majorUnit.HasValue)
                {
                    targetAxis.MajorUnitIsAuto = false;
                    targetAxis.MajorUnit = majorUnit.Value;
                }

                if (minorUnit.HasValue)
                {
                    targetAxis.MinorUnitIsAuto = false;
                    targetAxis.MinorUnit = minorUnit.Value;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref targetAxis!);
                ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    // === GRIDLINES ===

    /// <inheritdoc />
    public GridlinesResult GetGridlines(IExcelBatch batch, string chartName)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? axes = null;
            dynamic? valueAxis = null;
            dynamic? categoryAxis = null;

            try
            {
                axes = findResult.Chart.Axes;

                var result = new GridlinesResult
                {
                    Success = true,
                    ChartName = chartName,
                    Gridlines = new GridlinesInfo()
                };

                // Get value axis (type 2) gridlines
                try
                {
                    valueAxis = axes.Item(2); // xlValue
                    result.Gridlines.HasValueMajorGridlines = valueAxis.HasMajorGridlines;
                    result.Gridlines.HasValueMinorGridlines = valueAxis.HasMinorGridlines;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Value axis may not exist for some chart types
                }

                // Get category axis (type 1) gridlines
                try
                {
                    categoryAxis = axes.Item(1); // xlCategory
                    result.Gridlines.HasCategoryMajorGridlines = categoryAxis.HasMajorGridlines;
                    result.Gridlines.HasCategoryMinorGridlines = categoryAxis.HasMinorGridlines;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Category axis may not exist for some chart types
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref categoryAxis!);
                ComUtilities.Release(ref valueAxis!);
                ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetGridlines(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        bool? showMajor = null,
        bool? showMinor = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? axes = null;
            dynamic? targetAxis = null;

            try
            {
                axes = findResult.Chart.Axes;

                // Map axis type to Excel constants
                var (axisType, axisGroup) = MapAxisType(axis);
                targetAxis = axes.Item(axisType, axisGroup);

                if (showMajor.HasValue)
                    targetAxis.HasMajorGridlines = showMajor.Value;

                if (showMinor.HasValue)
                    targetAxis.HasMinorGridlines = showMinor.Value;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref targetAxis!);
                ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    // === SERIES FORMATTING ===

    /// <inheritdoc />
    public OperationResult SetSeriesFormat(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
        MarkerStyle? markerStyle = null,
        int? markerSize = null,
        string? markerBackgroundColor = null,
        string? markerForegroundColor = null,
        bool? invertIfNegative = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? seriesCollection = null;
            dynamic? series = null;

            try
            {
                seriesCollection = findResult.Chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                if (seriesIndex < 1 || seriesIndex > seriesCount)
                {
                    throw new ArgumentException($"Series index {seriesIndex} is out of range. Chart has {seriesCount} series (1-based indexing, use 1 for first series).");
                }

                series = seriesCollection.Item(seriesIndex);

                // Set marker style
                if (markerStyle.HasValue)
                    series.MarkerStyle = (int)markerStyle.Value;

                // Set marker size (valid range: 2-72)
                if (markerSize.HasValue)
                {
                    if (markerSize.Value < 2 || markerSize.Value > 72)
                    {
                        throw new ArgumentException($"Marker size must be between 2 and 72. Provided: {markerSize.Value}");
                    }
                    series.MarkerSize = markerSize.Value;
                }

                // Set marker background color (fill)
                if (!string.IsNullOrEmpty(markerBackgroundColor))
                {
                    int bgColor = ParseHexColor(markerBackgroundColor);
                    series.MarkerBackgroundColor = bgColor;
                }

                // Set marker foreground color (border)
                if (!string.IsNullOrEmpty(markerForegroundColor))
                {
                    int fgColor = ParseHexColor(markerForegroundColor);
                    series.MarkerForegroundColor = fgColor;
                }

                // Set invert if negative
                if (invertIfNegative.HasValue)
                    series.InvertIfNegative = invertIfNegative.Value;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref series!);
                ComUtilities.Release(ref seriesCollection!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    // === HELPER METHODS ===

    /// <summary>
    /// Maps ChartAxisType to Excel axis type and axis group constants.
    /// </summary>
    private static (int axisType, int axisGroup) MapAxisType(ChartAxisType axis)
    {
        return axis switch
        {
            ChartAxisType.Category => (1, 1),           // xlCategory, xlPrimary
            ChartAxisType.Value => (2, 1),              // xlValue, xlPrimary
            ChartAxisType.Primary => (1, 1),            // xlCategory, xlPrimary
            ChartAxisType.Secondary => (2, 1),          // xlValue, xlPrimary
            ChartAxisType.CategorySecondary => (1, 2),  // xlCategory, xlSecondary
            ChartAxisType.ValueSecondary => (2, 2),     // xlValue, xlSecondary
            _ => (1, 1)
        };
    }

    /// <summary>
    /// Parses a hex color string (#RRGGBB) to an Excel color integer (BGR format).
    /// </summary>
    private static int ParseHexColor(string hexColor)
    {
        // Remove # prefix if present
        string colorValue = hexColor.TrimStart('#');

        if (colorValue.Length != 6)
        {
            throw new ArgumentException($"Invalid hex color format: {hexColor}. Use #RRGGBB format.");
        }

        // Parse RGB components
        int r = Convert.ToInt32(colorValue[..2], 16);
        int g = Convert.ToInt32(colorValue.Substring(2, 2), 16);
        int b = Convert.ToInt32(colorValue.Substring(4, 2), 16);

        // Excel uses BGR format (Blue * 65536 + Green * 256 + Red)
        return b * 65536 + g * 256 + r;
    }

    // === TRENDLINE OPERATIONS ===

    /// <inheritdoc />
    public TrendlineListResult ListTrendlines(IExcelBatch batch, string chartName, int seriesIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? seriesCollection = null;
            dynamic? series = null;
            dynamic? trendlines = null;

            try
            {
                seriesCollection = findResult.Chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                if (seriesIndex < 1 || seriesIndex > seriesCount)
                {
                    throw new ArgumentException($"Series index {seriesIndex} is out of range. Chart has {seriesCount} series (1-based indexing, use 1 for first series).");
                }

                series = seriesCollection.Item(seriesIndex);
                trendlines = series.Trendlines();

                var result = new TrendlineListResult
                {
                    Success = true,
                    ChartName = chartName,
                    SeriesIndex = seriesIndex,
                    SeriesName = series.Name?.ToString() ?? $"Series {seriesIndex}"
                };

                int trendlineCount = trendlines.Count;
                for (int i = 1; i <= trendlineCount; i++)
                {
                    dynamic? trendline = null;
                    try
                    {
                        trendline = trendlines.Item(i);
                        var info = new TrendlineInfo
                        {
                            Index = i,
                            Type = (TrendlineType)Convert.ToInt32(trendline.Type),
                            Name = trendline.Name?.ToString(),
                            DisplayEquation = trendline.DisplayEquation,
                            DisplayRSquared = trendline.DisplayRSquared
                        };

                        // Get forward/backward forecast periods
                        try { info.Forward = trendline.Forward; } catch (COMException) { /* Optional COM property */ }
                        try { info.Backward = trendline.Backward; } catch (COMException) { /* Optional COM property */ }
                        try { info.Intercept = trendline.Intercept; } catch (COMException) { /* Optional COM property */ }

                        // Get order for polynomial trendlines
                        if (info.Type == TrendlineType.Polynomial)
                        {
                            try { info.Order = Convert.ToInt32(trendline.Order); } catch (COMException) { /* Optional COM property */ }
                        }

                        // Get period for moving average
                        if (info.Type == TrendlineType.MovingAverage)
                        {
                            try { info.Period = Convert.ToInt32(trendline.Period); } catch (COMException) { /* Optional COM property */ }
                        }

                        result.Trendlines.Add(info);
                    }
                    finally
                    {
                        ComUtilities.Release(ref trendline!);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref trendlines!);
                ComUtilities.Release(ref series!);
                ComUtilities.Release(ref seriesCollection!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public TrendlineResult AddTrendline(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
        TrendlineType trendlineType,
        int? order = null,
        int? period = null,
        double? forward = null,
        double? backward = null,
        double? intercept = null,
        bool displayEquation = false,
        bool displayRSquared = false,
        string? name = null)
    {
        // Validate type-specific parameters
        if (trendlineType == TrendlineType.Polynomial && (!order.HasValue || order.Value < 2 || order.Value > 6))
        {
            throw new ArgumentException("Polynomial trendline requires order parameter (2-6).");
        }

        if (trendlineType == TrendlineType.MovingAverage && (!period.HasValue || period.Value < 2))
        {
            throw new ArgumentException("Moving average trendline requires period parameter (2 or greater).");
        }

        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? seriesCollection = null;
            dynamic? series = null;
            dynamic? trendlines = null;
            dynamic? newTrendline = null;

            try
            {
                seriesCollection = findResult.Chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                if (seriesIndex < 1 || seriesIndex > seriesCount)
                {
                    throw new ArgumentException($"Series index {seriesIndex} is out of range. Chart has {seriesCount} series (1-based indexing, use 1 for first series).");
                }

                series = seriesCollection.Item(seriesIndex);
                trendlines = series.Trendlines();

                // Add trendline with type
                newTrendline = trendlines.Add((int)trendlineType);

                // Set optional parameters
                if (order.HasValue && trendlineType == TrendlineType.Polynomial)
                {
                    newTrendline.Order = order.Value;
                }

                if (period.HasValue && trendlineType == TrendlineType.MovingAverage)
                {
                    newTrendline.Period = period.Value;
                }

                if (forward.HasValue)
                {
                    newTrendline.Forward = forward.Value;
                }

                if (backward.HasValue)
                {
                    newTrendline.Backward = backward.Value;
                }

                if (intercept.HasValue)
                {
                    newTrendline.Intercept = intercept.Value;
                }

                newTrendline.DisplayEquation = displayEquation;
                newTrendline.DisplayRSquared = displayRSquared;

                if (!string.IsNullOrEmpty(name))
                {
                    newTrendline.Name = name;
                }

                // Get the index of the newly added trendline
                int trendlineIndex = trendlines.Count;

                return new TrendlineResult
                {
                    Success = true,
                    ChartName = chartName,
                    SeriesIndex = seriesIndex,
                    TrendlineIndex = trendlineIndex,
                    Type = trendlineType,
                    Name = name
                };
            }
            finally
            {
                ComUtilities.Release(ref newTrendline!);
                ComUtilities.Release(ref trendlines!);
                ComUtilities.Release(ref series!);
                ComUtilities.Release(ref seriesCollection!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteTrendline(IExcelBatch batch, string chartName, int seriesIndex, int trendlineIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? seriesCollection = null;
            dynamic? series = null;
            dynamic? trendlines = null;
            dynamic? trendline = null;

            try
            {
                seriesCollection = findResult.Chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                if (seriesIndex < 1 || seriesIndex > seriesCount)
                {
                    throw new ArgumentException($"Series index {seriesIndex} is out of range. Chart has {seriesCount} series (1-based indexing, use 1 for first series).");
                }

                series = seriesCollection.Item(seriesIndex);
                trendlines = series.Trendlines();
                int trendlineCount = trendlines.Count;

                if (trendlineIndex < 1 || trendlineIndex > trendlineCount)
                {
                    throw new ArgumentException($"Trendline index {trendlineIndex} is out of range. Series has {trendlineCount} trendlines.");
                }

                trendline = trendlines.Item(trendlineIndex);
                trendline.Delete();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref trendline!);
                ComUtilities.Release(ref trendlines!);
                ComUtilities.Release(ref series!);
                ComUtilities.Release(ref seriesCollection!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetTrendline(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
        int trendlineIndex,
        double? forward = null,
        double? backward = null,
        double? intercept = null,
        bool? displayEquation = null,
        bool? displayRSquared = null,
        string? name = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? seriesCollection = null;
            dynamic? series = null;
            dynamic? trendlines = null;
            dynamic? trendline = null;

            try
            {
                seriesCollection = findResult.Chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                if (seriesIndex < 1 || seriesIndex > seriesCount)
                {
                    throw new ArgumentException($"Series index {seriesIndex} is out of range. Chart has {seriesCount} series (1-based indexing, use 1 for first series).");
                }

                series = seriesCollection.Item(seriesIndex);
                trendlines = series.Trendlines();
                int trendlineCount = trendlines.Count;

                if (trendlineIndex < 1 || trendlineIndex > trendlineCount)
                {
                    throw new ArgumentException($"Trendline index {trendlineIndex} is out of range. Series has {trendlineCount} trendlines.");
                }

                trendline = trendlines.Item(trendlineIndex);

                // Update optional parameters
                if (forward.HasValue)
                {
                    trendline.Forward = forward.Value;
                }

                if (backward.HasValue)
                {
                    trendline.Backward = backward.Value;
                }

                if (intercept.HasValue)
                {
                    trendline.Intercept = intercept.Value;
                }

                if (displayEquation.HasValue)
                {
                    trendline.DisplayEquation = displayEquation.Value;
                }

                if (displayRSquared.HasValue)
                {
                    trendline.DisplayRSquared = displayRSquared.Value;
                }

                if (!string.IsNullOrEmpty(name))
                {
                    trendline.Name = name;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref trendline!);
                ComUtilities.Release(ref trendlines!);
                ComUtilities.Release(ref series!);
                ComUtilities.Release(ref seriesCollection!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult FitToRange(IExcelBatch batch, string chartName, string sheetName, string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            dynamic? worksheet = null;
            dynamic? range = null;

            try
            {
                // Get the target range
                worksheet = ctx.Book.Worksheets[sheetName];
                range = worksheet.Range[rangeAddress];

                // Get range geometry
                double left = Convert.ToDouble(range.Left);
                double top = Convert.ToDouble(range.Top);
                double width = Convert.ToDouble(range.Width);
                double height = Convert.ToDouble(range.Height);

                // Apply to chart shape
                findResult.Shape.Left = left;
                findResult.Shape.Top = top;
                findResult.Shape.Width = width;
                findResult.Shape.Height = height;

                // Collision detection after repositioning
                var warnings = ChartPositionHelpers.DetectCollisions(
                    worksheet, left, top, width, height, chartName);
                int chartCount = ChartPositionHelpers.CountCharts(worksheet);

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Message = ChartPositionHelpers.FormatCollisionWarnings(warnings, chartCount)
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref worksheet!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }
}



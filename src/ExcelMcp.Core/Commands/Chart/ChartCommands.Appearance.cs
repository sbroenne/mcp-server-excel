using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart appearance operations - type, title, axes, legend, style.
/// </summary>
public partial class ChartCommands
{
    /// <inheritdoc />
    public void SetChartType(IExcelBatch batch, string chartName, ChartType chartType)
    {
        batch.Execute((ctx, ct) =>
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

                return 0; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public void SetTitle(IExcelBatch batch, string chartName, string title)
    {
        batch.Execute((ctx, ct) =>
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

                return 0; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public void SetAxisTitle(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        string title)
    {
        batch.Execute((ctx, ct) =>
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

                return 0; // Void operation completed
            }
            finally
            {
                if (targetAxis != null) ComUtilities.Release(ref targetAxis!);
                if (axes != null) ComUtilities.Release(ref axes!);
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
                if (tickLabels != null) ComUtilities.Release(ref tickLabels!);
                if (targetAxis != null) ComUtilities.Release(ref targetAxis!);
                if (axes != null) ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public void SetAxisNumberFormat(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        string numberFormat)
    {
        batch.Execute((ctx, ct) =>
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

                return 0; // Void operation completed
            }
            finally
            {
                if (tickLabels != null) ComUtilities.Release(ref tickLabels!);
                if (targetAxis != null) ComUtilities.Release(ref targetAxis!);
                if (axes != null) ComUtilities.Release(ref axes!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public void ShowLegend(
        IExcelBatch batch,
        string chartName,
        bool visible,
        LegendPosition? position = null)
    {
        batch.Execute((ctx, ct) =>
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
                if (visible && position.HasValue)
                {
                    legend = findResult.Chart.Legend;
                    legend.Position = (int)position.Value;
                }

                return 0; // Void operation completed
            }
            finally
            {
                if (legend != null) ComUtilities.Release(ref legend!);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public void SetStyle(IExcelBatch batch, string chartName, int styleId)
    {
        batch.Execute((ctx, ct) =>
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
                    throw new ArgumentException($"Chart style ID must be between 1 and 48. Provided: {styleId}", nameof(styleId));
                }

                // Set chart style
                findResult.Chart.ChartStyle = styleId;

                return 0; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }
}

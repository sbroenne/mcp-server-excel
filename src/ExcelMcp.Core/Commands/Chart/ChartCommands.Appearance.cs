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

                return new OperationResult { Success = true };
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

                return new OperationResult { Success = true };
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

                return new OperationResult { Success = true };
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
    public OperationResult ShowLegend(
        IExcelBatch batch,
        string chartName,
        bool visible,
        LegendPosition? position = null)
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
                if (visible && position.HasValue)
                {
                    legend = findResult.Chart.Legend;
                    legend.Position = (int)position.Value;
                }

                return new OperationResult { Success = true };
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
                    throw new ArgumentException($"Chart style ID must be between 1 and 48. Provided: {styleId}", nameof(styleId));
                }

                // Set chart style
                findResult.Chart.ChartStyle = styleId;

                return new OperationResult { Success = true };
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }
}

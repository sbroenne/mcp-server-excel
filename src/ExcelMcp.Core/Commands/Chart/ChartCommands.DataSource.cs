using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Result of finding a chart by name.
/// </summary>
internal sealed class ChartFindResult
{
    public dynamic? Chart { get; set; }
    public dynamic? Shape { get; set; }
    public string SheetName { get; set; } = string.Empty;
}

/// <summary>
/// Chart data source operations - set range, add/remove series.
/// </summary>
public partial class ChartCommands
{
    /// <inheritdoc />
    public void SetSourceRange(IExcelBatch batch, string chartName, string sourceRange)
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
                // Determine strategy and delegate
                IChartStrategy strategy = _pivotStrategy.CanHandle(findResult.Chart) ? _pivotStrategy : _regularStrategy;
#pragma warning disable CS8604 // CodeQL false positive: Both strategies implement IChartStrategy.SetSourceRange with dynamic parameter
                strategy.SetSourceRange(findResult.Chart, sourceRange);
#pragma warning restore CS8604

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
    public SeriesInfo AddSeries(
        IExcelBatch batch,
        string chartName,
        string seriesName,
        string valuesRange,
        string? categoryRange = null)
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
                // Determine strategy and delegate
                IChartStrategy strategy = _pivotStrategy.CanHandle(findResult.Chart) ? _pivotStrategy : _regularStrategy;
#pragma warning disable CS8604 // CodeQL false positive: Both strategies implement IChartStrategy.AddSeries with dynamic parameter
                var result = strategy.AddSeries(findResult.Chart, seriesName, valuesRange, categoryRange);
#pragma warning restore CS8604

                return result;
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public void RemoveSeries(IExcelBatch batch, string chartName, int seriesIndex)
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
                // Determine strategy and delegate
                IChartStrategy strategy = _pivotStrategy.CanHandle(findResult.Chart) ? _pivotStrategy : _regularStrategy;
#pragma warning disable CS8604 // CodeQL false positive: Both strategies implement IChartStrategy.RemoveSeries with dynamic parameter
                strategy.RemoveSeries(findResult.Chart, seriesIndex);
#pragma warning restore CS8604

                return 0; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <summary>
    /// Finds a chart by name across all worksheets.
    /// Returns result with chart, shape, sheetName properties. Caller must release chart and shape.
    /// </summary>
    private static ChartFindResult FindChart(dynamic workbook, string chartName)
    {
        dynamic worksheets = workbook.Worksheets;
        int wsCount = Convert.ToInt32(worksheets.Count);

        for (int i = 1; i <= wsCount; i++)
        {
            dynamic? worksheet = null;
            dynamic? shapes = null;

            try
            {
                worksheet = worksheets.Item(i);
                string sheetName = worksheet.Name?.ToString() ?? $"Sheet{i}";
                shapes = worksheet.Shapes;
                int shapeCount = Convert.ToInt32(shapes.Count);

                for (int j = 1; j <= shapeCount; j++)
                {
                    dynamic? shape = null;
                    dynamic? chart = null;

                    try
                    {
                        shape = shapes.Item(j);

                        // Check if this is a chart (msoChart = 3)
                        if (Convert.ToInt32(shape.Type) != 3)
                        {
                            ComUtilities.Release(ref shape!);
                            continue;
                        }

                        string shapeName = shape.Name?.ToString() ?? string.Empty;
                        if (!shapeName.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                        {
                            ComUtilities.Release(ref shape!);
                            continue;
                        }

                        // Found it!
                        chart = shape.Chart;
                        ComUtilities.Release(ref shapes!);
                        ComUtilities.Release(ref worksheet!);
                        ComUtilities.Release(ref worksheets!);

                        return new ChartFindResult { Chart = chart, Shape = shape, SheetName = sheetName }; // Caller must release both
                    }
                    catch
                    {
                        if (chart != null) ComUtilities.Release(ref chart!);
                        if (shape != null) ComUtilities.Release(ref shape!);
                        throw;
                    }
                }
            }
            finally
            {
                if (shapes != null) ComUtilities.Release(ref shapes!);
                if (worksheet != null) ComUtilities.Release(ref worksheet!);
            }
        }

        ComUtilities.Release(ref worksheets!);
        return new ChartFindResult { Chart = null, Shape = null, SheetName = string.Empty };
    }
}

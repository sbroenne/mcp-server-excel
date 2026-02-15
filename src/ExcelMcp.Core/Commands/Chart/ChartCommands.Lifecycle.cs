using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart lifecycle operations - list, read, create, delete, move/resize.
/// </summary>
public partial class ChartCommands : IChartCommands, IChartConfigCommands
{
    private readonly RegularChartStrategy _regularStrategy = new();
    private readonly PivotChartStrategy _pivotStrategy = new();

    /// <inheritdoc />
    public List<ChartInfo> List(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var charts = new List<ChartInfo>();

            dynamic worksheets = ctx.Book.Worksheets;
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
                                continue;
                            }

                            chart = shape.Chart;
                            string chartName = shape.Name?.ToString() ?? $"Chart{j}";

                            // Determine strategy and get info
                            IChartStrategy strategy = _pivotStrategy.CanHandle(chart) ? _pivotStrategy : _regularStrategy;
#pragma warning disable CS8604 // CodeQL false positive: Both strategies implement IChartStrategy.GetInfo with dynamic parameters
                            var chartInfo = strategy.GetInfo(chart, chartName, sheetName, shape);
#pragma warning restore CS8604

                            charts.Add(chartInfo);
                        }
                        finally
                        {
                            ComUtilities.Release(ref chart!);
                            ComUtilities.Release(ref shape!);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shapes!);
                    ComUtilities.Release(ref worksheet!);
                }
            }

            ComUtilities.Release(ref worksheets!);

            return charts;
        });
    }

    /// <inheritdoc />
    public ChartInfoResult Read(IExcelBatch batch, string chartName)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name across all worksheets
            dynamic worksheets = ctx.Book.Worksheets;
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

                            // Check if this is a chart and name matches
                            if (Convert.ToInt32(shape.Type) != 3)
                            {
                                continue;
                            }

                            string shapeName = shape.Name?.ToString() ?? string.Empty;
                            if (!shapeName.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            chart = shape.Chart;

                            // Determine strategy and get detailed info
                            IChartStrategy strategy = _pivotStrategy.CanHandle(chart) ? _pivotStrategy : _regularStrategy;
#pragma warning disable CS8604 // CodeQL false positive: Both strategies implement IChartStrategy.GetDetailedInfo with dynamic parameters
                            var result = strategy.GetDetailedInfo(chart, chartName, sheetName, shape);
#pragma warning restore CS8604

                            ComUtilities.Release(ref chart!);
                            ComUtilities.Release(ref shape!);
                            ComUtilities.Release(ref shapes!);
                            ComUtilities.Release(ref worksheet!);
                            ComUtilities.Release(ref worksheets!);

                            return result;
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            ComUtilities.Release(ref chart!);
                            ComUtilities.Release(ref shape!);
                            throw;
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shapes!);
                    ComUtilities.Release(ref worksheet!);
                }
            }

            ComUtilities.Release(ref worksheets!);

            // Chart not found
            throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
        });
    }

    /// <inheritdoc />
    public ChartCreateResult CreateFromRange(
        IExcelBatch batch,
        string sheetName,
        string sourceRange,
        ChartType chartType,
        double left,
        double top,
        double width = 400,
        double height = 300,
        string? chartName = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? worksheet = null;
            dynamic? shapes = null;
            dynamic? shape = null;
            dynamic? chart = null;

            try
            {
                worksheet = ctx.Book.Worksheets.Item(sheetName);
                shapes = worksheet.Shapes;

                // Create chart using AddChart
                shape = shapes.AddChart(
                    XlChartType: (int)chartType,
                    Left: left,
                    Top: top,
                    Width: width,
                    Height: height
                );

                chart = shape.Chart;

                // Set data source - need to get Range object from string address
                dynamic? sourceRangeObj = null;
                try
                {
                    // Get the range object from the address string
                    // If sourceRange doesn't include sheet name, prefix it
                    string fullRangeAddress = sourceRange.Contains('!')
                        ? sourceRange
                        : $"{sheetName}!{sourceRange}";
                    sourceRangeObj = ctx.Book.Application.Range(fullRangeAddress);
                    chart.SetSourceData(sourceRangeObj);
                }
                finally
                {
                    if (sourceRangeObj != null)
                    {
                        ComUtilities.Release(ref sourceRangeObj!);
                    }
                }

                // Set custom name if provided
                if (!string.IsNullOrWhiteSpace(chartName))
                {
                    shape.Name = chartName;
                }

                string finalName = shape.Name?.ToString() ?? "Chart";

                var result = new ChartCreateResult
                {
                    Success = true,
                    ChartName = finalName,
                    SheetName = sheetName,
                    ChartType = chartType,
                    IsPivotChart = false,
                    Left = left,
                    Top = top,
                    Width = width,
                    Height = height
                };

                return result;
            }
            finally
            {
                ComUtilities.Release(ref chart!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref worksheet!);
            }
        });
    }

    /// <inheritdoc />
    public ChartCreateResult CreateFromTable(
        IExcelBatch batch,
        string tableName,
        string sheetName,
        ChartType chartType,
        double left,
        double top,
        double width = 400,
        double height = 300,
        string? chartName = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? tableRange = null;
            dynamic? worksheet = null;
            dynamic? shapes = null;
            dynamic? shape = null;
            dynamic? chart = null;

            try
            {
                // Find the table using CoreLookupHelpers
                table = CoreLookupHelpers.FindTable(ctx.Book, tableName);

                // Get the table's data range (includes headers)
                tableRange = table.Range;

                // Get target worksheet
                worksheet = ctx.Book.Worksheets.Item(sheetName);
                shapes = worksheet.Shapes;

                // Create chart using AddChart
                shape = shapes.AddChart(
                    XlChartType: (int)chartType,
                    Left: left,
                    Top: top,
                    Width: width,
                    Height: height
                );

                chart = shape.Chart;

                // Set data source to table's range
                chart.SetSourceData(tableRange);

                // Set custom name if provided
                if (!string.IsNullOrWhiteSpace(chartName))
                {
                    shape.Name = chartName;
                }

                string finalName = shape.Name?.ToString() ?? "Chart";

                var result = new ChartCreateResult
                {
                    Success = true,
                    ChartName = finalName,
                    SheetName = sheetName,
                    ChartType = chartType,
                    IsPivotChart = false,
                    Left = left,
                    Top = top,
                    Width = width,
                    Height = height
                };

                return result;
            }
            finally
            {
                ComUtilities.Release(ref chart!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref worksheet!);
                ComUtilities.Release(ref tableRange!);
                ComUtilities.Release(ref table!);
            }
        });
    }

    /// <inheritdoc />
    public ChartCreateResult CreateFromPivotTable(
        IExcelBatch batch,
        string pivotTableName,
        string sheetName,
        ChartType chartType,
        double left,
        double top,
        double width = 400,
        double height = 300,
        string? chartName = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? worksheet = null;
            dynamic? pivotChartShape = null;
            dynamic? chart = null;
            dynamic? pivotTable = null;
            dynamic? tableRange = null;
            dynamic? shapes = null;

            try
            {
                // Find PivotTable
                pivotTable = FindPivotTable(ctx.Book, pivotTableName);
                if (pivotTable == null)
                {
                    throw new InvalidOperationException($"PivotTable '{pivotTableName}' not found in workbook.");
                }

                // Get target worksheet
                worksheet = ctx.Book.Worksheets.Item(sheetName);

                // Create a chart via Shapes.AddChart and set source to PivotTable's range.
                // This approach works for both OLAP (Data Model) and non-OLAP PivotTables,
                // unlike PivotCache.CreatePivotChart which has parameter issues in dynamic
                // COM and throws DISP_E_UNKNOWNNAME for OLAP sources.
                shapes = worksheet.Shapes;

                // Create chart using AddChart
                pivotChartShape = shapes.AddChart(
                    XlChartType: (int)chartType,
                    Left: left,
                    Top: top,
                    Width: width,
                    Height: height
                );

                chart = pivotChartShape.Chart;

                // Get the PivotTable's data range and set it as the chart's source
                tableRange = pivotTable.TableRange1;
                chart.SetSourceData(tableRange);

                // Set custom name if provided
                if (!string.IsNullOrWhiteSpace(chartName))
                {
                    pivotChartShape.Name = chartName;
                }

                string finalName = pivotChartShape.Name?.ToString() ?? "Chart";

                var result = new ChartCreateResult
                {
                    Success = true,
                    ChartName = finalName,
                    SheetName = sheetName,
                    ChartType = chartType,
                    IsPivotChart = true,
                    LinkedPivotTable = pivotTableName,
                    Left = left,
                    Top = top,
                    Width = width,
                    Height = height
                };

                return result;
            }
            finally
            {
                ComUtilities.Release(ref chart!);
                ComUtilities.Release(ref pivotChartShape!);
                ComUtilities.Release(ref tableRange!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref worksheet!);
                ComUtilities.Release(ref pivotTable!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string chartName)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find and delete chart
            dynamic worksheets = ctx.Book.Worksheets;
            int wsCount = Convert.ToInt32(worksheets.Count);

            for (int i = 1; i <= wsCount; i++)
            {
                dynamic? worksheet = null;
                dynamic? shapes = null;

                try
                {
                    worksheet = worksheets.Item(i);
                    shapes = worksheet.Shapes;
                    int shapeCount = Convert.ToInt32(shapes.Count);

                    for (int j = 1; j <= shapeCount; j++)
                    {
                        dynamic? shape = null;

                        try
                        {
                            shape = shapes.Item(j);

                            // Check if this is a chart and name matches
                            if (Convert.ToInt32(shape.Type) != 3)
                            {
                                continue;
                            }

                            string shapeName = shape.Name?.ToString() ?? string.Empty;
                            if (!shapeName.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            // Delete the chart
                            shape.Delete();

                            ComUtilities.Release(ref shape!);
                            ComUtilities.Release(ref shapes!);
                            ComUtilities.Release(ref worksheet!);
                            ComUtilities.Release(ref worksheets!);

                            return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Success
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
                    ComUtilities.Release(ref worksheet!);
                }
            }

            ComUtilities.Release(ref worksheets!);

            // Chart not found
            throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
        });
    }

    /// <inheritdoc />
    public OperationResult Move(
        IExcelBatch batch,
        string chartName,
        double? left = null,
        double? top = null,
        double? width = null,
        double? height = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart and update position/size
            dynamic worksheets = ctx.Book.Worksheets;
            int wsCount = Convert.ToInt32(worksheets.Count);

            for (int i = 1; i <= wsCount; i++)
            {
                dynamic? worksheet = null;
                dynamic? shapes = null;

                try
                {
                    worksheet = worksheets.Item(i);
                    shapes = worksheet.Shapes;
                    int shapeCount = Convert.ToInt32(shapes.Count);

                    for (int j = 1; j <= shapeCount; j++)
                    {
                        dynamic? shape = null;

                        try
                        {
                            shape = shapes.Item(j);

                            // Check if this is a chart and name matches
                            if (Convert.ToInt32(shape.Type) != 3)
                            {
                                continue;
                            }

                            string shapeName = shape.Name?.ToString() ?? string.Empty;
                            if (!shapeName.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            // Update position and size
                            if (left.HasValue) shape.Left = left.Value;
                            if (top.HasValue) shape.Top = top.Value;
                            if (width.HasValue) shape.Width = width.Value;
                            if (height.HasValue) shape.Height = height.Value;

                            ComUtilities.Release(ref shape!);
                            ComUtilities.Release(ref shapes!);
                            ComUtilities.Release(ref worksheet!);
                            ComUtilities.Release(ref worksheets!);

                            return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Success
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
                    ComUtilities.Release(ref worksheet!);
                }
            }

            ComUtilities.Release(ref worksheets!);

            // Chart not found
            throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
        });
    }

    /// <summary>
    /// Finds a PivotTable by name across all worksheets.
    /// Delegates to CoreLookupHelpers.TryFindPivotTable for the actual lookup.
    /// </summary>
    private static dynamic? FindPivotTable(dynamic workbook, string pivotTableName)
    {
        CoreLookupHelpers.TryFindPivotTable(workbook, pivotTableName, out dynamic? pivotTable);
        return pivotTable;
    }
}



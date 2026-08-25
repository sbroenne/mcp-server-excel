using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
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
    public ChartListResult List(IExcelBatch batch)
    {
        var result = new ChartListResult
        {
            Action = "list",
            FilePath = batch.WorkbookPath
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
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

                                result.Charts.Add(chartInfo);
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

                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref worksheets!);
            }
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
        string sourceRangeAddress,
        ChartType chartType,
        double left = 0,
        double top = 0,
        double width = 400,
        double height = 300,
        string? chartName = null,
        string? targetRange = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? worksheet = null;
            dynamic? shapes = null;
            dynamic? shape = null;
            dynamic? chart = null;
            dynamic? targetRangeObj = null;

            try
            {
                worksheet = ctx.Book.Worksheets[sheetName];
                shapes = worksheet.Shapes;

                // Resolve final position: targetRange > explicit left/top > auto-position
                double finalLeft = left;
                double finalTop = top;
                double finalWidth = width;
                double finalHeight = height;

                if (!string.IsNullOrWhiteSpace(targetRange))
                {
                    // targetRange takes precedence — resolve range geometry
                    targetRangeObj = worksheet.Range[targetRange];
                    finalLeft = Convert.ToDouble(targetRangeObj.Left);
                    finalTop = Convert.ToDouble(targetRangeObj.Top);
                    finalWidth = Convert.ToDouble(targetRangeObj.Width);
                    finalHeight = Convert.ToDouble(targetRangeObj.Height);
                }
                else if (left == 0 && top == 0)
                {
                    // No explicit position — auto-position below content
                    // Cast explicitly to avoid dynamic dispatch losing named tuple members
                    (double Left, double Top) autoPos = ChartPositionHelpers.FindAvailablePosition(worksheet, width, height);
                    finalLeft = autoPos.Left;
                    finalTop = autoPos.Top;
                }

                // Create chart using AddChart
                shape = shapes.AddChart(
                    XlChartType: (int)chartType,
                    Left: finalLeft,
                    Top: finalTop,
                    Width: finalWidth,
                    Height: finalHeight
                );

                chart = shape.Chart;

                // Set data source - need to get Range object from string address
                dynamic? sourceRangeObj = null;
                try
                {
                    // Get the range object from the address string
                    // If sourceRangeAddress doesn't include sheet name, prefix it
                    // Sheet names with spaces or special characters must be quoted: 'Sheet Name'!A1:D6
                    string fullRangeAddress = sourceRangeAddress.Contains('!')
                        ? sourceRangeAddress
                        : $"'{sheetName}'!{sourceRangeAddress}";
                    sourceRangeObj = ctx.Book.Application.Range[fullRangeAddress];
                    try
                    {
                        chart.SetSourceData(sourceRangeObj);
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                        when (ex.HResult == unchecked((int)0x800A03EC))
                    {
                        throw new InvalidOperationException(
                            $"Cannot set chart data source to '{sourceRangeAddress}'. " +
                            "The range must be contiguous, non-empty, and accessible. " +
                            "If the data is not in a table, consider creating a table first with " +
                            "table(action='create'), then use chart(action='create-from-table').", ex);
                    }
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

                // Collision detection — warn about overlaps after positioning
                var warnings = ChartPositionHelpers.DetectCollisions(
                    worksheet, finalLeft, finalTop, finalWidth, finalHeight, finalName);
                int chartCount = ChartPositionHelpers.CountCharts(worksheet);

                var result = new ChartCreateResult
                {
                    Success = true,
                    ChartName = finalName,
                    SheetName = sheetName,
                    ChartType = chartType,
                    IsPivotChart = false,
                    Left = finalLeft,
                    Top = finalTop,
                    Width = finalWidth,
                    Height = finalHeight,
                    Message = ChartPositionHelpers.FormatCollisionWarnings(warnings, chartCount)
                };

                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetRangeObj!);
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
        double left = 0,
        double top = 0,
        double width = 400,
        double height = 300,
        string? chartName = null,
        string? targetRange = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? tableRange = null;
            dynamic? worksheet = null;
            dynamic? shapes = null;
            dynamic? shape = null;
            dynamic? chart = null;
            dynamic? targetRangeObj = null;

            try
            {
                // Find the table using CoreLookupHelpers
                table = CoreLookupHelpers.FindTable(ctx.Book, tableName);

                // Get the table's data range (includes headers)
                tableRange = table.Range;

                // Get target worksheet
                worksheet = ctx.Book.Worksheets[sheetName];
                shapes = worksheet.Shapes;

                // Resolve final position: targetRange > explicit left/top > auto-position
                double finalLeft = left;
                double finalTop = top;
                double finalWidth = width;
                double finalHeight = height;

                if (!string.IsNullOrWhiteSpace(targetRange))
                {
                    targetRangeObj = worksheet.Range[targetRange];
                    finalLeft = Convert.ToDouble(targetRangeObj.Left);
                    finalTop = Convert.ToDouble(targetRangeObj.Top);
                    finalWidth = Convert.ToDouble(targetRangeObj.Width);
                    finalHeight = Convert.ToDouble(targetRangeObj.Height);
                }
                else if (left == 0 && top == 0)
                {
                    // Cast explicitly to avoid dynamic dispatch losing named tuple members
                    (double Left, double Top) autoPos = ChartPositionHelpers.FindAvailablePosition(worksheet, width, height);
                    finalLeft = autoPos.Left;
                    finalTop = autoPos.Top;
                }

                // Create chart using AddChart
                shape = shapes.AddChart(
                    XlChartType: (int)chartType,
                    Left: finalLeft,
                    Top: finalTop,
                    Width: finalWidth,
                    Height: finalHeight
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

                // Collision detection
                var warnings = ChartPositionHelpers.DetectCollisions(
                    worksheet, finalLeft, finalTop, finalWidth, finalHeight, finalName);
                int chartCount = ChartPositionHelpers.CountCharts(worksheet);

                var result = new ChartCreateResult
                {
                    Success = true,
                    ChartName = finalName,
                    SheetName = sheetName,
                    ChartType = chartType,
                    IsPivotChart = false,
                    Left = finalLeft,
                    Top = finalTop,
                    Width = finalWidth,
                    Height = finalHeight,
                    Message = ChartPositionHelpers.FormatCollisionWarnings(warnings, chartCount)
                };

                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetRangeObj!);
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
        double left = 0,
        double top = 0,
        double width = 400,
        double height = 300,
        string? chartName = null,
        string? targetRange = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? worksheet = null;
            Excel.Worksheet? sourceWorksheet = null;
            Excel.Worksheet? previousWorksheet = null;
            Excel.Range? previousCell = null;
            Excel.Shape? pivotChartShape = null;
            Excel.Chart? chart = null;
            Excel.PivotTable? pivotTable = null;
            Excel.Range? tableRange = null;
            Excel.Sheets? worksheets = null;
            Excel.Shapes? shapes = null;
            Excel.Range? targetRangeObj = null;
            Excel.PivotLayout? pivotLayout = null;
            Excel.PivotTable? linkedPivotTable = null;

            try
            {
                // Find PivotTable
                pivotTable = (Excel.PivotTable?)FindPivotTable(ctx.Book, pivotTableName);
                if (pivotTable == null)
                {
                    throw new InvalidOperationException($"PivotTable '{pivotTableName}' not found in workbook.");
                }

                // Get target worksheet
                worksheets = (Excel.Sheets)ctx.Book.Worksheets;
                worksheet = (Excel.Worksheet)worksheets[sheetName];
                sourceWorksheet = (Excel.Worksheet)pivotTable.Parent;
                tableRange = pivotTable.TableRange1;
                previousWorksheet = ctx.App.ActiveSheet as Excel.Worksheet;
                previousCell = ctx.App.ActiveCell as Excel.Range;

                // Resolve final position: targetRange > explicit left/top > auto-position
                double finalLeft = left;
                double finalTop = top;
                double finalWidth = width;
                double finalHeight = height;

                if (!string.IsNullOrWhiteSpace(targetRange))
                {
                    targetRangeObj = worksheet.Range[targetRange];
                    finalLeft = targetRangeObj.Left;
                    finalTop = targetRangeObj.Top;
                    finalWidth = targetRangeObj.Width;
                    finalHeight = targetRangeObj.Height;
                }
                else if (left == 0 && top == 0)
                {
                    // Cast explicitly to avoid dynamic dispatch losing named tuple members
                    (double Left, double Top) autoPos = ChartPositionHelpers.FindAvailablePosition(worksheet, width, height);
                    finalLeft = autoPos.Left;
                    finalTop = autoPos.Top;
                }

                // Match Excel's Insert PivotChart sequence: activate/select the source
                // PivotTable before adding the chart, then bind the selected PivotTable range.
                sourceWorksheet.Activate();
                tableRange.Select();
                shapes = worksheet.Shapes;

                try
                {
                    pivotChartShape = shapes.AddChart2(
                        Style: -1,
                        XlChartType: (int)chartType,
                        Left: finalLeft,
                        Top: finalTop,
                        Width: finalWidth,
                        Height: finalHeight,
                        NewLayout: true);

                    chart = pivotChartShape.Chart;
                    chart.SetSourceData(tableRange);
                }
                catch (COMException ex)
                {
                    DeleteCreatedChart(pivotChartShape);
                    throw new InvalidOperationException(
                        $"Excel could not create a linked PivotChart from PivotTable '{pivotTableName}' " +
                        $"with chart type '{chartType}'. The requested chart type or PivotTable layout " +
                        "is not supported for PivotCharts.", ex);
                }

                try
                {
                    pivotLayout = chart.PivotLayout;
                    linkedPivotTable = pivotLayout?.PivotTable;
                }
                catch (COMException ex)
                {
                    DeleteCreatedChart(pivotChartShape);
                    throw new InvalidOperationException(
                        $"Excel created a regular chart instead of a linked PivotChart for PivotTable " +
                        $"'{pivotTableName}'. No chart was kept.", ex);
                }

                string linkedPivotTableName = linkedPivotTable?.Name ?? string.Empty;
                if (!linkedPivotTableName.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                {
                    DeleteCreatedChart(pivotChartShape);
                    throw new InvalidOperationException(
                        $"Excel did not link the new PivotChart to PivotTable '{pivotTableName}'. " +
                        "No chart was kept.");
                }

                // Set custom name if provided
                if (!string.IsNullOrWhiteSpace(chartName))
                {
                    pivotChartShape.Name = chartName;
                }

                string finalName = pivotChartShape.Name?.ToString() ?? "Chart";

                // Collision detection
                var warnings = ChartPositionHelpers.DetectCollisions(
                    worksheet, finalLeft, finalTop, finalWidth, finalHeight, finalName);
                int chartCount = ChartPositionHelpers.CountCharts(worksheet);

                var result = new ChartCreateResult
                {
                    Success = true,
                    ChartName = finalName,
                    SheetName = sheetName,
                    ChartType = (ChartType)chart.ChartType,
                    IsPivotChart = true,
                    LinkedPivotTable = linkedPivotTableName,
                    Left = finalLeft,
                    Top = finalTop,
                    Width = finalWidth,
                    Height = finalHeight,
                    Message = ChartPositionHelpers.FormatCollisionWarnings(warnings, chartCount)
                };

                return result;
            }
            finally
            {
                if (previousWorksheet != null)
                {
                    try
                    {
                        previousWorksheet.Activate();
                        previousCell?.Select();
                    }
                    catch (COMException)
                    {
                        // Restoring the prior selection is best-effort after Excel completed the operation.
                    }
                }

                ComUtilities.Release(ref linkedPivotTable);
                ComUtilities.Release(ref pivotLayout);
                ComUtilities.Release(ref targetRangeObj!);
                ComUtilities.Release(ref chart!);
                ComUtilities.Release(ref pivotChartShape!);
                ComUtilities.Release(ref tableRange!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref previousCell);
                ComUtilities.Release(ref previousWorksheet);
                ComUtilities.Release(ref sourceWorksheet);
                ComUtilities.Release(ref worksheet!);
                ComUtilities.Release(ref pivotTable!);
            }
        });
    }

    private static void DeleteCreatedChart(Excel.Shape? chartShape)
    {
        if (chartShape == null)
        {
            return;
        }

        try
        {
            chartShape.Delete();
        }
        catch (COMException)
        {
            // Preserve the creation failure if Excel already removed the partial chart.
        }
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

                            // Collision detection after repositioning
                            double finalLeft = Convert.ToDouble(shape.Left);
                            double finalTop = Convert.ToDouble(shape.Top);
                            double finalWidth = Convert.ToDouble(shape.Width);
                            double finalHeight = Convert.ToDouble(shape.Height);

                            var warnings = ChartPositionHelpers.DetectCollisions(
                                worksheet, finalLeft, finalTop, finalWidth, finalHeight, shapeName);
                            int chartCount = ChartPositionHelpers.CountCharts(worksheet);

                            ComUtilities.Release(ref shape!);
                            ComUtilities.Release(ref shapes!);
                            ComUtilities.Release(ref worksheet!);
                            ComUtilities.Release(ref worksheets!);

                            return new OperationResult
                            {
                                Success = true,
                                FilePath = batch.WorkbookPath,
                                Message = ChartPositionHelpers.FormatCollisionWarnings(warnings, chartCount)
                            };
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

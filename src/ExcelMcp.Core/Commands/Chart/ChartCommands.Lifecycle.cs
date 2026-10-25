using Sbroenne.ExcelMcp.Core.Utilities;
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
            Excel.Sheets? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                int wsCount = Convert.ToInt32(worksheets.Count);

                for (int i = 1; i <= wsCount; i++)
                {
                    Excel.Worksheet? worksheet = null;
                    Excel.Shapes? shapes = null;

                    try
                    {
                        worksheet = (Excel.Worksheet)worksheets.Item[i];
                        string sheetName = worksheet.Name?.ToString() ?? $"Sheet{i}";
                        shapes = worksheet.Shapes;
                        int shapeCount = Convert.ToInt32(shapes.Count);

                        for (int j = 1; j <= shapeCount; j++)
                        {
                            Excel.Shape? shape = null;
                            Excel.Chart? chart = null;

                            try
                            {
                                shape = shapes.Item(j);

                                // Check if this is a chart (msoChart = 3)
                                if ((int)shape.Type != 3)
                                {
                                    continue;
                                }

                                chart = shape.Chart;
                                string chartName = shape.Name?.ToString() ?? $"Chart{j}";

                                // Determine strategy and get info
                                IChartStrategy strategy = _pivotStrategy.CanHandle(chart) ? _pivotStrategy : _regularStrategy;
                                var chartInfo = strategy.GetInfo(chart, chartName, sheetName, shape);

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
            Excel.Sheets? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                int wsCount = Convert.ToInt32(worksheets.Count);

                for (int i = 1; i <= wsCount; i++)
                {
                    Excel.Worksheet? worksheet = null;
                    Excel.Shapes? shapes = null;

                    try
                    {
                        worksheet = (Excel.Worksheet)worksheets.Item[i];
                        string sheetName = worksheet.Name?.ToString() ?? $"Sheet{i}";
                        shapes = worksheet.Shapes;
                        int shapeCount = Convert.ToInt32(shapes.Count);

                        for (int j = 1; j <= shapeCount; j++)
                        {
                            Excel.Shape? shape = null;
                            Excel.Chart? chart = null;

                            try
                            {
                                shape = shapes.Item(j);

                                if ((int)shape.Type != 3)
                                {
                                    continue;
                                }

                                string shapeName = shape.Name?.ToString() ?? string.Empty;
                                if (!shapeName.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                chart = shape.Chart;
                                IChartStrategy strategy = _pivotStrategy.CanHandle(chart) ? _pivotStrategy : _regularStrategy;
                                return strategy.GetDetailedInfo(chart, chartName, sheetName, shape);
                            }
                            finally
                            {
                                ComUtilities.Release(ref chart);
                                ComUtilities.Release(ref shape);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref shapes);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
            }
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
            Excel.Worksheet? worksheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.Chart? chart = null;
            Excel.Range? targetRangeObj = null;
            Excel.Range? sourceRangeObj = null;

            try
            {
                worksheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);
                shapes = worksheet.Shapes;

                // Resolve the source before creating anything so a bad address leaves no chart behind.
                sourceRangeObj = ChartSourceRange.Resolve(ctx.Book, sheetName, sourceRangeAddress, ct);

                // Resolve final position: targetRange > explicit left/top > auto-position
                double finalLeft = left;
                double finalTop = top;
                double finalWidth = width;
                double finalHeight = height;

                if (!string.IsNullOrWhiteSpace(targetRange))
                {
                    // targetRange takes precedence — resolve range geometry
                    targetRangeObj = worksheet.Range[targetRange];
                    finalLeft = Convert.ToDouble((object)targetRangeObj.Left, System.Globalization.CultureInfo.InvariantCulture);
                    finalTop = Convert.ToDouble((object)targetRangeObj.Top, System.Globalization.CultureInfo.InvariantCulture);
                    finalWidth = Convert.ToDouble((object)targetRangeObj.Width, System.Globalization.CultureInfo.InvariantCulture);
                    finalHeight = Convert.ToDouble((object)targetRangeObj.Height, System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (left == 0 && top == 0)
                {
                    // No explicit position — auto-position below content
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
                string finalName = ConfigureCreatedChart(
                    shape, chart, sourceRangeObj, sourceRangeAddress, sheetName, chartName);

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
                ComUtilities.Release(ref sourceRangeObj!);
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
            Excel.ListObject? table = null;
            Excel.Range? tableRange = null;
            Excel.Worksheet? worksheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.Chart? chart = null;
            Excel.Range? targetRangeObj = null;

            try
            {
                // Find the table using CoreLookupHelpers
                table = CoreLookupHelpers.FindTable(ctx.Book, tableName);

                // Get the table's data range (includes headers)
                tableRange = table.Range;

                // Get target worksheet
                worksheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);
                shapes = worksheet.Shapes;

                // Resolve final position: targetRange > explicit left/top > auto-position
                double finalLeft = left;
                double finalTop = top;
                double finalWidth = width;
                double finalHeight = height;

                if (!string.IsNullOrWhiteSpace(targetRange))
                {
                    targetRangeObj = worksheet.Range[targetRange];
                    finalLeft = Convert.ToDouble((object)targetRangeObj.Left, System.Globalization.CultureInfo.InvariantCulture);
                    finalTop = Convert.ToDouble((object)targetRangeObj.Top, System.Globalization.CultureInfo.InvariantCulture);
                    finalWidth = Convert.ToDouble((object)targetRangeObj.Width, System.Globalization.CultureInfo.InvariantCulture);
                    finalHeight = Convert.ToDouble((object)targetRangeObj.Height, System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (left == 0 && top == 0)
                {
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
                string finalName = ConfigureCreatedChart(
                    shape, chart, tableRange, tableName, sheetName, chartName);

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
            string? createdChartName = null;
            string creationStep = $"Creating a linked PivotChart for '{pivotTableName}'";
            Exception? operationFailure = null;

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
                    finalLeft = Convert.ToDouble((object)targetRangeObj.Left, System.Globalization.CultureInfo.InvariantCulture);
                    finalTop = Convert.ToDouble((object)targetRangeObj.Top, System.Globalization.CultureInfo.InvariantCulture);
                    finalWidth = Convert.ToDouble((object)targetRangeObj.Width, System.Globalization.CultureInfo.InvariantCulture);
                    finalHeight = Convert.ToDouble((object)targetRangeObj.Height, System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (left == 0 && top == 0)
                {
                    (double Left, double Top) autoPos = ChartPositionHelpers.FindAvailablePosition(worksheet, width, height);
                    finalLeft = autoPos.Left;
                    finalTop = autoPos.Top;
                }

                // Match Excel's Insert PivotChart sequence: activate/select the source
                // PivotTable before adding the chart, then bind the selected PivotTable range.
                sourceWorksheet.Activate();
                tableRange.Select();
                shapes = worksheet.Shapes;

                pivotChartShape = shapes.AddChart2(
                    Style: -1,
                    XlChartType: (int)chartType,
                    Left: finalLeft,
                    Top: finalTop,
                    Width: finalWidth,
                    Height: finalHeight,
                    NewLayout: true);
                createdChartName = "(name unavailable)";
                creationStep = "Reading the new chart's identity";
                createdChartName = pivotChartShape.Name ?? createdChartName;

                creationStep = $"Binding the new chart to PivotTable '{pivotTableName}'";
                chart = pivotChartShape.Chart;
                chart.SetSourceData(tableRange);

                creationStep = $"Verifying the PivotChart link to '{pivotTableName}'";
                pivotLayout = chart.PivotLayout;
                linkedPivotTable = pivotLayout?.PivotTable;

                string linkedPivotTableName = linkedPivotTable?.Name ?? string.Empty;
                if (!linkedPivotTableName.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"Excel did not link the new PivotChart to PivotTable '{pivotTableName}'.");
                }

                // Set custom name if provided
                if (!string.IsNullOrWhiteSpace(chartName))
                {
                    creationStep = $"Renaming the new chart to '{chartName}'";
                    pivotChartShape.Name = chartName;
                    createdChartName = chartName;
                }

                creationStep = "Reading the created PivotChart's details and collision warnings";
                string finalName = pivotChartShape.Name?.ToString() ?? "Chart";
                createdChartName = finalName;

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
            catch (Exception ex) when (createdChartName != null && CreatedObjectFailure.CanReport(ex))
            {
                operationFailure = DescribePivotChartCreationFailure(createdChartName, sheetName, creationStep, ex);
                throw operationFailure;
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                try
                {
                    if (previousWorksheet != null)
                    {
                        OperationCleanup.Run(() =>
                        {
                            previousWorksheet.Activate();
                            previousCell?.Select();
                        }, operationFailure);
                    }
                }
                finally
                {
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
            }
        });
    }

    private static string ConfigureCreatedChart(
        Excel.Shape shape,
        Excel.Chart chart,
        Excel.Range sourceRange,
        string sourceDescription,
        string sheetName,
        string? chartName)
    {
        string createdName = shape.Name?.ToString() ?? "Chart";
        string step = $"Setting the data source of the new chart to '{sourceDescription}'";
        try
        {
            ChartSourceRange.Apply(chart, sourceRange, sourceDescription);
            if (!string.IsNullOrWhiteSpace(chartName))
            {
                step = $"Renaming the new chart to '{chartName}'";
                shape.Name = chartName;
            }

            return shape.Name?.ToString() ?? createdName;
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
        {
            throw CreatedObjectFailure.Create("chart", createdName, sheetName, step, ex);
        }
    }

    internal static InvalidOperationException DescribePivotChartCreationFailure(
        string createdName, string sheetName, string step, Exception cause) =>
        CreatedObjectFailure.Create("chart", createdName, sheetName, step, cause);

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string chartName)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find and delete chart
            Excel.Sheets? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                int wsCount = Convert.ToInt32(worksheets.Count);

                for (int i = 1; i <= wsCount; i++)
                {
                    Excel.Worksheet? worksheet = null;
                    Excel.Shapes? shapes = null;

                    try
                    {
                        worksheet = (Excel.Worksheet)worksheets.Item[i];
                        shapes = worksheet.Shapes;
                        int shapeCount = Convert.ToInt32(shapes.Count);

                        for (int j = 1; j <= shapeCount; j++)
                        {
                            Excel.Shape? shape = null;

                            try
                            {
                                shape = shapes.Item(j);

                                if ((int)shape.Type != 3)
                                {
                                    continue;
                                }

                                string shapeName = shape.Name?.ToString() ?? string.Empty;
                                if (!shapeName.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                shape.Delete();
                                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
                            }
                            finally
                            {
                                ComUtilities.Release(ref shape);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref shapes);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
            }
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
            Excel.Sheets? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                int wsCount = worksheets.Count;
                for (int i = 1; i <= wsCount; i++)
                {
                    Excel.Worksheet? worksheet = null;
                    Excel.Shapes? shapes = null;

                    try
                    {
                        worksheet = (Excel.Worksheet)worksheets.Item[i];
                        shapes = worksheet.Shapes;
                        int shapeCount = shapes.Count;
                        for (int j = 1; j <= shapeCount; j++)
                        {
                            Excel.Shape? shape = null;
                            try
                            {
                                shape = shapes.Item(j);
                                if ((int)shape.Type != 3)
                                {
                                    continue;
                                }

                                string shapeName = shape.Name?.ToString() ?? string.Empty;
                                if (!shapeName.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                if (left.HasValue) shape.Left = (float)left.Value;
                                if (top.HasValue) shape.Top = (float)top.Value;
                                if (width.HasValue) shape.Width = (float)width.Value;
                                if (height.HasValue) shape.Height = (float)height.Value;

                                double finalLeft = shape.Left;
                                double finalTop = shape.Top;
                                double finalWidth = shape.Width;
                                double finalHeight = shape.Height;
                                var warnings = ChartPositionHelpers.DetectCollisions(
                                    worksheet, finalLeft, finalTop, finalWidth, finalHeight, shapeName);
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
                                ComUtilities.Release(ref shape);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref shapes);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
            }
        });
    }

    /// <summary>
    /// Finds a PivotTable by name across all worksheets.
    /// Delegates to CoreLookupHelpers.TryFindPivotTable for the actual lookup.
    /// </summary>
    private static Excel.PivotTable? FindPivotTable(Excel.Workbook workbook, string pivotTableName)
    {
        CoreLookupHelpers.TryFindPivotTable(workbook, pivotTableName, out Excel.PivotTable? pivotTable);
        return pivotTable;
    }
}

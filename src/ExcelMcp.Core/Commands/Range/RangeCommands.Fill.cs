using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult Fill(IExcelBatch batch, string sheetName, string rangeAddress,
        FillDirection direction, OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty)
    {
        if (!Enum.IsDefined(direction))
            throw new ArgumentOutOfRangeException(nameof(direction));
        ValidateOverwritePolicy(overwritePolicy);
        return ExecuteFill(batch, sheetName, rangeAddress, "fill", (ctx, range, ct) =>
        {
            var size = GetContentDimensions(range);
            EnsureDestinationWritable(ctx, range, overwritePolicy, ct, (row, column) => direction switch
            {
                FillDirection.Down => row > 0,
                FillDirection.Up => row < size.Rows - 1,
                FillDirection.Left => column < size.Columns - 1,
                FillDirection.Right => column > 0,
                _ => throw new ArgumentOutOfRangeException(nameof(direction))
            });
            ct.ThrowIfCancellationRequested();
            switch (direction)
            {
                case FillDirection.Down: range.FillDown(); break;
                case FillDirection.Up: range.FillUp(); break;
                case FillDirection.Left: range.FillLeft(); break;
                case FillDirection.Right: range.FillRight(); break;
                default: throw new ArgumentOutOfRangeException(nameof(direction));
            }
        });
    }

    /// <inheritdoc />
    public OperationResult AutoFill(IExcelBatch batch, string sheetName, string sourceRange,
        string destinationRange, AutoFillKind fillType = AutoFillKind.Default,
        OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty)
    {
        if (!Enum.IsDefined(fillType))
            throw new ArgumentOutOfRangeException(nameof(fillType));
        ValidateOverwritePolicy(overwritePolicy);
        bool replacesSource = fillType is AutoFillKind.LinearTrend or AutoFillKind.GrowthTrend;
        if (replacesSource && overwritePolicy != OverwritePolicy.Allow)
            throw new OperationFailureException(OperationFailureCategory.Conflict,
                "Trend AutoFill can replace source values. Use overwrite_policy='allow' " +
                "(CLI: --overwrite-policy allow) only for authorized replacement.");
        return ExecuteFill(batch, sheetName, destinationRange, "auto-fill", (ctx, destination, ct) =>
        {
            Excel.Range? source = null;
            Excel.Worksheet? sourceSheet = null;
            Excel.Worksheet? destinationSheet = null;
            try
            {
                source = ResolveFillRange(ctx, sheetName, sourceRange);
                var sourceSize = GetContentDimensions(source);
                var destinationSize = GetContentDimensions(destination);
                sourceSheet = source.Worksheet;
                destinationSheet = destination.Worksheet;
                if (!string.Equals(sourceSheet.Name, destinationSheet.Name, StringComparison.Ordinal))
                    throw new ArgumentException("AutoFill source and destination must be on the same worksheet.");

                int rowOffset = source.Row - destination.Row;
                int columnOffset = source.Column - destination.Column;
                bool vertical = columnOffset == 0 && sourceSize.Columns == destinationSize.Columns &&
                    destinationSize.Rows > sourceSize.Rows &&
                    (rowOffset == 0 || rowOffset == destinationSize.Rows - sourceSize.Rows);
                bool horizontal = rowOffset == 0 && sourceSize.Rows == destinationSize.Rows &&
                    destinationSize.Columns > sourceSize.Columns &&
                    (columnOffset == 0 || columnOffset == destinationSize.Columns - sourceSize.Columns);
                if (!vertical && !horizontal)
                    throw new ArgumentException(
                        "AutoFill destination must include the source and extend it in exactly one direction.");
                if (fillType != AutoFillKind.Formats)
                {
                    EnsureDestinationWritable(ctx, destination, overwritePolicy, ct, (row, column) =>
                        replacesSource || row < rowOffset || row >= rowOffset + sourceSize.Rows ||
                        column < columnOffset || column >= columnOffset + sourceSize.Columns);
                }
                ct.ThrowIfCancellationRequested();
                source.AutoFill(destination, fillType switch
                {
                    AutoFillKind.Default => Excel.XlAutoFillType.xlFillDefault,
                    AutoFillKind.Copy => Excel.XlAutoFillType.xlFillCopy,
                    AutoFillKind.Series => Excel.XlAutoFillType.xlFillSeries,
                    AutoFillKind.Formats => Excel.XlAutoFillType.xlFillFormats,
                    AutoFillKind.WithoutFormatting => Excel.XlAutoFillType.xlFillValues,
                    AutoFillKind.Days => Excel.XlAutoFillType.xlFillDays,
                    AutoFillKind.Weekdays => Excel.XlAutoFillType.xlFillWeekdays,
                    AutoFillKind.Months => Excel.XlAutoFillType.xlFillMonths,
                    AutoFillKind.Years => Excel.XlAutoFillType.xlFillYears,
                    AutoFillKind.LinearTrend => Excel.XlAutoFillType.xlLinearTrend,
                    AutoFillKind.GrowthTrend => Excel.XlAutoFillType.xlGrowthTrend,
                    AutoFillKind.FlashFill => Excel.XlAutoFillType.xlFlashFill,
                    _ => throw new ArgumentOutOfRangeException(nameof(fillType))
                });
            }
            finally
            {
                ComUtilities.Release(ref destinationSheet);
                ComUtilities.Release(ref sourceSheet);
                ComUtilities.Release(ref source);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult CreateSeries(IExcelBatch batch, string sheetName, string rangeAddress,
        SeriesOrientation orientation, SeriesKind seriesType = SeriesKind.Linear, double stepValue = 1,
        double? stopValue = null, SeriesDateUnit dateUnit = SeriesDateUnit.Day, bool trend = false,
        OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty)
    {
        if (!Enum.IsDefined(orientation))
            throw new ArgumentOutOfRangeException(nameof(orientation));
        if (!Enum.IsDefined(seriesType))
            throw new ArgumentOutOfRangeException(nameof(seriesType));
        if (!Enum.IsDefined(dateUnit))
            throw new ArgumentOutOfRangeException(nameof(dateUnit));
        if (!double.IsFinite(stepValue) || stepValue == 0)
            throw new ArgumentOutOfRangeException(nameof(stepValue), "Series step must be finite and nonzero.");
        if (stopValue.HasValue && !double.IsFinite(stopValue.Value))
            throw new ArgumentOutOfRangeException(nameof(stopValue), "Series stop must be finite.");
        if (trend && seriesType is not (SeriesKind.Linear or SeriesKind.Growth))
            throw new ArgumentException("Trend fitting requires a linear or growth series.", nameof(trend));
        ValidateOverwritePolicy(overwritePolicy);
        if (trend && overwritePolicy != OverwritePolicy.Allow)
            throw new OperationFailureException(OperationFailureCategory.Conflict,
                "Series trend fitting can replace input values. Use overwrite_policy='allow' " +
                "(CLI: --overwrite-policy allow) only for authorized replacement.");
        return ExecuteFill(batch, sheetName, rangeAddress, "create-series", (ctx, range, ct) =>
        {
            EnsureDestinationWritable(ctx, range, overwritePolicy, ct, (row, column) =>
                trend || (orientation == SeriesOrientation.Columns ? row > 0 : column > 0));
            ct.ThrowIfCancellationRequested();
            range.DataSeries(
                orientation == SeriesOrientation.Rows ? Excel.XlRowCol.xlRows : Excel.XlRowCol.xlColumns,
                seriesType switch
                {
                    SeriesKind.Linear => Excel.XlDataSeriesType.xlDataSeriesLinear,
                    SeriesKind.Growth => Excel.XlDataSeriesType.xlGrowth,
                    SeriesKind.Date => Excel.XlDataSeriesType.xlChronological,
                    SeriesKind.AutoFill => Excel.XlDataSeriesType.xlAutoFill,
                    _ => throw new ArgumentOutOfRangeException(nameof(seriesType))
                },
                dateUnit switch
                {
                    SeriesDateUnit.Day => Excel.XlDataSeriesDate.xlDay,
                    SeriesDateUnit.Weekday => Excel.XlDataSeriesDate.xlWeekday,
                    SeriesDateUnit.Month => Excel.XlDataSeriesDate.xlMonth,
                    SeriesDateUnit.Year => Excel.XlDataSeriesDate.xlYear,
                    _ => throw new ArgumentOutOfRangeException(nameof(dateUnit))
                },
                stepValue, stopValue.HasValue ? stopValue.Value : Type.Missing, trend);
        });
    }

    private static OperationResult ExecuteFill(IExcelBatch batch, string sheetName, string rangeAddress,
        string action, Action<ExcelContext, Excel.Range, CancellationToken> fill)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                range = ResolveFillRange(ctx, sheetName, rangeAddress);
                fill(ctx, range, ct);
                return new OperationResult { FilePath = batch.WorkbookPath, Action = action, Success = true };
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    private static Excel.Range ResolveFillRange(ExcelContext context, string sheetName, string rangeAddress)
    {
        Excel.Range? range = null;
        try
        {
            range = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
            if (range is null)
                throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
            GetContentDimensions(range);
            if (RangeMergeDiscovery.GetMergeCellsState(range.MergeCells) != false)
                throw new OperationFailureException(OperationFailureCategory.Conflict,
                    "Native filling requires an unmerged rectangle. No write was attempted.");
            var resolved = range;
            range = null;
            return resolved;
        }
        finally
        {
            ComUtilities.Release(ref range);
        }
    }
}

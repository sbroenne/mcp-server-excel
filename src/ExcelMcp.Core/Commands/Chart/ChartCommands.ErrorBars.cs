using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

public partial class ChartCommands
{
    /// <inheritdoc />
    public ChartErrorBarsResult GetErrorBars(IExcelBatch batch, string chartName, int seriesIndex)
    {
        ValidateSeriesIndex(seriesIndex);
        return WithNativeSeries(batch, chartName, seriesIndex, false,
            (_, series, _, _, _) => ReadErrorBars(batch, chartName, seriesIndex, series));
    }

    /// <inheritdoc />
    public ChartErrorBarsResult SetErrorBars(IExcelBatch batch, string chartName, int seriesIndex, ChartErrorBarOptions errorBarOptions)
    {
        ValidateSeriesIndex(seriesIndex);
        ValidateErrorOptions(errorBarOptions);
        return WithNativeSeries(batch, chartName, seriesIndex, true, (_, series, book, sheet, ct) =>
        {
            Excel.Range? plus = null;
            Excel.Range? minus = null;
            Excel.Points? points = null;
            Excel.ErrorBars? bars = null;
            try
            {
                if (!errorBarOptions.Enabled)
                {
                    series.HasErrorBars = false;
                    return ReadErrorBars(batch, chartName, seriesIndex, series);
                }
                if (errorBarOptions.Direction == ChartErrorBarDirection.X &&
                    !series.ChartType.ToString().Contains("Scatter", StringComparison.Ordinal) &&
                    !series.ChartType.ToString().Contains("Bubble", StringComparison.Ordinal))
                    throw new ArgumentException("X-direction error bars require an XY scatter or bubble series.");
                if (errorBarOptions.Kind == ChartErrorBarKind.Custom)
                {
                    points = (Excel.Points)series.Points();
                    var sourceSheet = errorBarOptions.SourceSheetName ?? sheet;
                    plus = (Excel.Range?)RangeHelpers.ResolveRange(book, sourceSheet, errorBarOptions.PlusRange!)
                        ?? throw new ArgumentException("Custom positive error range could not be resolved.");
                    minus = (Excel.Range?)RangeHelpers.ResolveRange(book, sourceSheet, errorBarOptions.MinusRange!)
                        ?? throw new ArgumentException("Custom negative error range could not be resolved.");
                    ValidateCustomErrorRange(plus, points.Count, ct);
                    ValidateCustomErrorRange(minus, points.Count, ct);
                }
                series.ErrorBar((Excel.XlErrorBarDirection)errorBarOptions.Direction,
                    (Excel.XlErrorBarInclude)errorBarOptions.Include, (Excel.XlErrorBarType)errorBarOptions.Kind,
                    errorBarOptions.Kind == ChartErrorBarKind.Custom ? plus : errorBarOptions.Amount ?? Type.Missing,
                    errorBarOptions.Kind == ChartErrorBarKind.Custom ? minus : Type.Missing);
                if (!series.HasErrorBars)
                    throw new InvalidOperationException("Excel did not create error bars for this series/chart type.");
                if (errorBarOptions.EndStyle.HasValue)
                {
                    bars = series.ErrorBars;
                    bars.EndStyle = (Excel.XlEndStyleCap)errorBarOptions.EndStyle.Value;
                }
                return ReadErrorBars(batch, chartName, seriesIndex, series);
            }
            finally
            {
                ComUtilities.Release(ref bars);
                ComUtilities.Release(ref points);
                ComUtilities.Release(ref minus);
                ComUtilities.Release(ref plus);
            }
        });
    }

    private static void ValidateErrorOptions(ChartErrorBarOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (!Enum.IsDefined(options.Kind) || !Enum.IsDefined(options.Direction) || !Enum.IsDefined(options.Include) ||
            (options.EndStyle.HasValue && !Enum.IsDefined(options.EndStyle.Value)))
            throw new ArgumentException("Unknown native error-bar enum value.", nameof(options));
        if (!options.Enabled)
        {
            if (options.Amount.HasValue || options.PlusRange != null || options.MinusRange != null || options.SourceSheetName != null || options.EndStyle.HasValue)
                throw new ArgumentException("Disabling error bars cannot be combined with calculation/range/end-style changes.", nameof(options));
            return;
        }
        if (options.Kind == ChartErrorBarKind.Custom)
        {
            if (string.IsNullOrWhiteSpace(options.PlusRange) || string.IsNullOrWhiteSpace(options.MinusRange) || options.Amount.HasValue ||
                (options.SourceSheetName != null && string.IsNullOrWhiteSpace(options.SourceSheetName)))
                throw new ArgumentException("Custom bars require plusRange and minusRange, with no scalar amount.", nameof(options));
        }
        else
        {
            if (options.PlusRange != null || options.MinusRange != null || options.SourceSheetName != null)
                throw new ArgumentException("Worksheet ranges/sourceSheetName are only valid for Custom error bars.", nameof(options));
            if (options.Kind == ChartErrorBarKind.StandardError)
            {
                if (options.Amount.HasValue)
                    throw new ArgumentException("StandardError calculates its own amount; omit amount.", nameof(options));
            }
            else if (!options.Amount.HasValue || !double.IsFinite(options.Amount.Value) || options.Amount < 0 ||
                (options.Kind == ChartErrorBarKind.StandardDeviation && options.Amount == 0))
                throw new ArgumentException("Supply a nonnegative fixed/percent amount or positive standard-deviation multiplier.", nameof(options));
        }
    }

    private static void ValidateCustomErrorRange(Excel.Range range, int count, CancellationToken ct)
    {
        Excel.Areas? areas = null;
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            areas = range.Areas;
            rows = range.Rows;
            columns = range.Columns;
            if (areas.Count != 1 || (rows.Count != 1 && columns.Count != 1) || Convert.ToDouble(range.CountLarge) != count)
                throw new ArgumentException("Each custom error range must be one contiguous row or column with one value per series point.");
            var values = range.Value2;
            var items = values is object[,] matrix ? matrix.Cast<object?>() : [values];
            foreach (var item in items)
            {
                ct.ThrowIfCancellationRequested();
                if (item is not double number || !double.IsFinite(number) || number < 0)
                    throw new ArgumentException("Custom error ranges require nonnegative numeric cells; blanks, text, logical values and errors are invalid.");
            }
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref areas);
        }
    }

    private static ChartErrorBarsResult ReadErrorBars(IExcelBatch batch, string name, int index, Excel.Series series)
    {
        var result = new ChartErrorBarsResult
        {
            Success = true,
            FilePath = batch.WorkbookPath,
            ChartName = name,
            SeriesIndex = index,
            HasErrorBars = series.HasErrorBars,
            SettingsReadable = false
        };
        if (!result.HasErrorBars) return result;
        Excel.ErrorBars? bars = null;
        try
        {
            bars = series.ErrorBars;
            result.EndStyle = (ChartErrorBarEndStyle)bars.EndStyle;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref bars);
        }
    }
}

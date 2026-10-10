using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

public sealed partial class SlicerCommands
{
    /// <inheritdoc />
    public SlicerStateResult CreateTimeline(IExcelBatch batch, string pivotTableName, string fieldName,
        string slicerName, string destinationSheet, string position)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? anchor = null;
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.Slicers? slicers = null;
            Excel.Slicer? slicer = null;
            try
            {
                pivot = CoreLookupHelpers.FindPivotTable(ctx.Book, pivotTableName);
                (sheet, anchor) = SlicerPlacement.ResolveDestination(ctx.Book, destinationSheet, position, ct);
                if (sheet.ProtectDrawingObjects)
                    throw new InvalidOperationException("Unprotect drawing objects before creating a timeline.");
                caches = ctx.Book.SlicerCaches;
                SlicerPlacement.ValidateNewControlName(caches, slicerName, ct);
                ct.ThrowIfCancellationRequested();
                cache = caches.Add2(pivot, fieldName, Type.Missing, Excel.XlSlicerCacheType.xlTimeline);
                slicers = cache.Slicers;
                try
                {
                    slicer = slicers.Add(sheet, Type.Missing, slicerName, slicerName, anchor.Top, anchor.Left);
                }
                catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
                {
                    throw SlicerPlacement.LeftoverCache(cache, fieldName, ex);
                }
                return State(batch, slicer, cache, ct);
            }
            finally
            {
                ComUtilities.Release(ref slicer);
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
                ComUtilities.Release(ref anchor);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <inheritdoc />
    public SlicerStateResult GetSlicer(IExcelBatch batch, string slicerName) =>
        WithSlicer(batch, slicerName, (_, slicer, cache, ct) => State(batch, slicer, cache, ct));

    /// <inheritdoc />
    public SlicerStateResult UpdateSlicer(IExcelBatch batch, string slicerName, SlicerUpdateOptions slicerOptions)
    {
        ArgumentNullException.ThrowIfNull(slicerOptions);
        foreach (var number in new[] { slicerOptions.Left, slicerOptions.Top, slicerOptions.Width, slicerOptions.Height })
            if (number.HasValue && (!double.IsFinite(number.Value) || number.Value < 0))
                throw new ArgumentException("Slicer coordinates and dimensions must be finite nonnegative point values.");
        if (slicerOptions.Width is <= 0 || slicerOptions.Height is <= 0 || slicerOptions.ColumnCount is <= 0)
            throw new ArgumentException("Slicer width, height and column count must be positive.");
        if (slicerOptions.Granularity.HasValue && !Enum.IsDefined(slicerOptions.Granularity.Value))
            throw new ArgumentException("Unknown timeline display granularity.");
        return WithSlicer(batch, slicerName, (book, slicer, cache, ct) =>
        {
            bool timeline = cache.SlicerCacheType == Excel.XlSlicerCacheType.xlTimeline;
            if (timeline && (slicerOptions.ColumnCount.HasValue || slicerOptions.DisplayHeader.HasValue))
                throw new ArgumentException("Timeline controls use timeline view settings, not ordinary column/header layout.");
            if (!timeline && (slicerOptions.Granularity.HasValue || slicerOptions.ShowHeader.HasValue ||
                slicerOptions.ShowSelectionLabel.HasValue || slicerOptions.ShowTimeLevel.HasValue ||
                slicerOptions.ShowHorizontalScrollbar.HasValue))
                throw new ArgumentException("Timeline view settings require a date timeline.");
            Excel.TableStyles? styles = null;
            Excel.TableStyle? style = null;
            Excel.TimelineViewState? view = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = (Excel.Worksheet)slicer.Parent;
                if (sheet.ProtectDrawingObjects)
                    throw new InvalidOperationException("Unprotect drawing objects before updating a control.");
                if (slicerOptions.Style is not null)
                {
                    styles = book.TableStyles;
                    style = styles[slicerOptions.Style];
                    if (timeline ? !style.ShowAsAvailableTimelineStyle : !style.ShowAsAvailableSlicerStyle)
                        throw new ArgumentException("The selected style is not available for this control type.");
                }
                if (timeline)
                    view = slicer.TimelineViewState;
                ct.ThrowIfCancellationRequested();
                if (slicerOptions.Left.HasValue) slicer.Left = slicerOptions.Left.Value;
                if (slicerOptions.Top.HasValue) slicer.Top = slicerOptions.Top.Value;
                if (slicerOptions.Width.HasValue) slicer.Width = slicerOptions.Width.Value;
                if (slicerOptions.Height.HasValue) slicer.Height = slicerOptions.Height.Value;
                if (slicerOptions.Caption is not null) slicer.Caption = slicerOptions.Caption;
                if (slicerOptions.Style is not null) slicer.Style = slicerOptions.Style;
                if (slicerOptions.ColumnCount.HasValue) slicer.NumberOfColumns = slicerOptions.ColumnCount.Value;
                if (slicerOptions.DisplayHeader.HasValue) slicer.DisplayHeader = slicerOptions.DisplayHeader.Value;
                if (view is not null)
                {
                    if (slicerOptions.Granularity.HasValue) view.Level = ToNativeLevel(slicerOptions.Granularity.Value);
                    if (slicerOptions.ShowHeader.HasValue) view.ShowHeader = slicerOptions.ShowHeader.Value;
                    if (slicerOptions.ShowSelectionLabel.HasValue) view.ShowSelectionLabel = slicerOptions.ShowSelectionLabel.Value;
                    if (slicerOptions.ShowTimeLevel.HasValue) view.ShowTimeLevel = slicerOptions.ShowTimeLevel.Value;
                    if (slicerOptions.ShowHorizontalScrollbar.HasValue) view.ShowHorizontalScrollbar = slicerOptions.ShowHorizontalScrollbar.Value;
                }
                return State(batch, slicer, cache, ct);
            }
            finally
            {
                ComUtilities.Release(ref view);
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SlicerStateResult SetTimelineSelection(IExcelBatch batch, string slicerName, TimelineSelectionOptions timelineSelection)
    {
        ArgumentNullException.ThrowIfNull(timelineSelection);
        if (timelineSelection.StartDate > timelineSelection.EndDate ||
            timelineSelection.StartDate.TimeOfDay != TimeSpan.Zero || timelineSelection.EndDate.TimeOfDay != TimeSpan.Zero)
            throw new ArgumentException("Timeline selection requires ascending calendar dates with no time components.");
        return WithSlicer(batch, slicerName, (_, slicer, cache, ct) =>
        {
            RequireTimeline(cache);
            Excel.TimelineState? state = null;
            try
            {
                state = cache.TimelineState;
                state.SetFilterDateRange(timelineSelection.StartDate, timelineSelection.EndDate);
                return State(batch, slicer, cache, ct);
            }
            finally
            {
                ComUtilities.Release(ref state);
            }
        });
    }

    /// <inheritdoc />
    public SlicerStateResult ClearTimelineSelection(IExcelBatch batch, string slicerName) =>
        WithSlicer(batch, slicerName, (_, slicer, cache, ct) =>
        {
            RequireTimeline(cache);
            cache.ClearDateFilter();
            return State(batch, slicer, cache, ct);
        });

    /// <inheritdoc />
    public SlicerStateResult ConnectPivotTable(IExcelBatch batch, string slicerName, string pivotTableName) =>
        ChangeConnection(batch, slicerName, pivotTableName, true);

    /// <inheritdoc />
    public SlicerStateResult DisconnectPivotTable(IExcelBatch batch, string slicerName, string pivotTableName) =>
        ChangeConnection(batch, slicerName, pivotTableName, false);

    private static SlicerStateResult ChangeConnection(IExcelBatch batch, string slicerName, string pivotTableName, bool connect) =>
        WithSlicer(batch, slicerName, (book, slicer, cache, ct) =>
        {
            if (cache.List)
                throw new ArgumentException("Excel Table slicers cannot connect to PivotTables.");
            Excel.PivotTable? target = null;
            Excel.SlicerPivotTables? pivots = null;
            Excel.PivotTable? source = null;
            try
            {
                target = CoreLookupHelpers.FindPivotTable(book, pivotTableName);
                pivots = cache.PivotTables;
                var names = ReadConnections(cache, ct);
                bool connected = names.Contains(target.Name, StringComparer.OrdinalIgnoreCase);
                if (connect)
                {
                    if (pivots.Count == 0)
                        throw new InvalidOperationException("The control has no source PivotTable connection.");
                    source = pivots[1];
                    if (source.CacheIndex != target.CacheIndex)
                        throw new ArgumentException("PivotTables must share the existing native PivotCache; no cache is rebuilt or replaced.");
                    if (!connected)
                        pivots.AddPivotTable(target);
                }
                else
                {
                    if (!connected)
                        throw new ArgumentException($"PivotTable '{pivotTableName}' is not connected to this control.");
                    if (pivots.Count == 1)
                        throw new ArgumentException("Keep at least one PivotTable connection; delete the control instead to remove it.");
                    pivots.RemovePivotTable(target);
                }
                return State(batch, slicer, cache, ct);
            }
            finally
            {
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref target);
            }
        });

    private static SlicerStateResult WithSlicer(IExcelBatch batch, string name,
        Func<Excel.Workbook, Excel.Slicer, Excel.SlicerCache, CancellationToken, SlicerStateResult> action)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.SlicerCaches? caches = null;
            try
            {
                caches = ctx.Book.SlicerCaches;
                for (int index = 1; index <= caches.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.SlicerCache? cache = null;
                    Excel.Slicers? slicers = null;
                    try
                    {
                        cache = caches[index];
                        slicers = cache.Slicers;
                        for (int item = 1; item <= slicers.Count; item++)
                        {
                            ct.ThrowIfCancellationRequested();
                            Excel.Slicer? slicer = null;
                            try
                            {
                                slicer = slicers[item];
                                if (string.Equals(slicer.Name, name, StringComparison.OrdinalIgnoreCase))
                                    return action(ctx.Book, slicer, cache, ct);
                            }
                            finally
                            {
                                ComUtilities.Release(ref slicer);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref slicers);
                        ComUtilities.Release(ref cache);
                    }
                }
                throw new ArgumentException($"Slicer or timeline '{name}' not found.");
            }
            finally
            {
                ComUtilities.Release(ref caches);
            }
        });
    }

    private static SlicerStateResult State(IExcelBatch batch, Excel.Slicer slicer, Excel.SlicerCache cache, CancellationToken ct)
    {
        var result = new SlicerStateResult { FilePath = batch.WorkbookPath };
        var details = result.Slicer;
        details.Name = slicer.Name;
        details.CacheName = cache.Name;
        details.FieldName = cache.SourceName;
        details.Caption = slicer.Caption;
        details.Left = slicer.Left;
        details.Top = slicer.Top;
        details.Width = slicer.Width;
        details.Height = slicer.Height;
        details.IsTimeline = cache.SlicerCacheType == Excel.XlSlicerCacheType.xlTimeline;
        details.IsTable = cache.List;
        details.FilterCleared = cache.FilterCleared;
        details.ConnectedPivotTables = ReadConnections(cache, ct);
        Excel.Worksheet? parent = null;
        Excel.ListObject? table = null;
        object? style = null;
        try
        {
            parent = (Excel.Worksheet)slicer.Parent;
            details.SheetName = parent.Name;
            style = slicer.Style;
            details.Style = style switch
            {
                string name => name,
                Excel.TableStyle native => native.Name,
                _ => throw new InvalidOperationException($"Unsupported native slicer style type '{style?.GetType().Name}'.")
            };
            if (details.IsTable)
            {
                table = cache.ListObject;
                details.ConnectedTable = table.Name;
            }
            if (details.IsTimeline)
                details.Timeline = ReadTimelineDetails(slicer, cache);
            else
            {
                details.ColumnCount = slicer.NumberOfColumns;
                details.DisplayHeader = slicer.DisplayHeader;
                ReadItems(slicer, cache, details, ct);
            }
            result.Success = true;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref style);
            ComUtilities.Release(ref parent);
        }
    }

    internal static TimelineDetails ReadTimelineDetails(Excel.Slicer slicer, Excel.SlicerCache cache)
    {
        Excel.TimelineViewState? view = null;
        Excel.TimelineState? state = null;
        try
        {
            view = slicer.TimelineViewState;
            var details = new TimelineDetails
            {
                Granularity = view.Level switch
                {
                    Excel.XlTimelineLevel.xlTimelineLevelYears => TimelineGranularity.Years,
                    Excel.XlTimelineLevel.xlTimelineLevelQuarters => TimelineGranularity.Quarters,
                    Excel.XlTimelineLevel.xlTimelineLevelMonths => TimelineGranularity.Months,
                    Excel.XlTimelineLevel.xlTimelineLevelDays => TimelineGranularity.Days,
                    _ => throw new InvalidOperationException("Unknown native timeline display level.")
                },
                ShowHeader = view.ShowHeader,
                ShowSelectionLabel = view.ShowSelectionLabel,
                ShowTimeLevel = view.ShowTimeLevel,
                ShowHorizontalScrollbar = view.ShowHorizontalScrollbar
            };
            if (!cache.FilterCleared)
            {
                state = cache.TimelineState;
                details.StartDate = Convert.ToDateTime(state.StartDate, CultureInfo.InvariantCulture);
                details.EndDate = Convert.ToDateTime(state.EndDate, CultureInfo.InvariantCulture);
                details.FilterType = state.FilterType.ToString();
                details.SingleRangeFilterState = state.SingleRangeFilterState;
            }
            return details;
        }
        finally
        {
            ComUtilities.Release(ref state);
            ComUtilities.Release(ref view);
        }
    }

    private static void ReadItems(Excel.Slicer slicer, Excel.SlicerCache cache, SlicerDetails details, CancellationToken ct)
    {
        Excel.SlicerCacheLevel? level = null;
        Excel.SlicerItems? items = null;
        try
        {
            if (cache.OLAP)
            {
                level = slicer.SlicerCacheLevel;
                items = level.SlicerItems;
            }
            else
                items = cache.SlicerItems;
            for (int index = 1; index <= items.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.SlicerItem? item = null;
                try
                {
                    item = items[index];
                    details.AvailableItems.Add(item.Name);
                    if (item.Selected)
                        details.SelectedItems.Add(item.Name);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref items);
            ComUtilities.Release(ref level);
        }
    }

    private static List<string> ReadConnections(Excel.SlicerCache cache, CancellationToken ct)
    {
        List<string> names = [];
        if (cache.List)
            return names;
        Excel.SlicerPivotTables? pivots = null;
        try
        {
            pivots = cache.PivotTables;
            for (int index = 1; index <= pivots.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotTable? pivot = null;
                try
                {
                    pivot = pivots[index];
                    names.Add(pivot.Name);
                }
                finally
                {
                    ComUtilities.Release(ref pivot);
                }
            }
            return names;
        }
        finally
        {
            ComUtilities.Release(ref pivots);
        }
    }

    private static void RequireTimeline(Excel.SlicerCache cache)
    {
        if (cache.SlicerCacheType != Excel.XlSlicerCacheType.xlTimeline)
            throw new ArgumentException("This operation requires a date timeline, not an ordinary slicer.");
    }

    private static Excel.XlTimelineLevel ToNativeLevel(TimelineGranularity level) => level switch
    {
        TimelineGranularity.Years => Excel.XlTimelineLevel.xlTimelineLevelYears,
        TimelineGranularity.Quarters => Excel.XlTimelineLevel.xlTimelineLevelQuarters,
        TimelineGranularity.Months => Excel.XlTimelineLevel.xlTimelineLevelMonths,
        TimelineGranularity.Days => Excel.XlTimelineLevel.xlTimelineLevelDays,
        _ => throw new ArgumentException("Unknown timeline display granularity.")
    };
}

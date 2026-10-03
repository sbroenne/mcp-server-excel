using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Filtering;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult ApplyFilter(IExcelBatch batch, string sheetName, string rangeAddress,
        int columnIndex, FilterOptions filterOptions)
    {
        NativeFilterHelpers.Validate(filterOptions);
        ArgumentOutOfRangeException.ThrowIfLessThan(columnIndex, 1);
        return batch.Execute((ctx, token) =>
        {
            Excel.Range? range = null;
            Excel.Worksheet? sheet = null;
            Excel.AutoFilter? filter = null;
            try
            {
                range = ResolveFillRange(ctx, sheetName, rangeAddress);
                var size = GetContentDimensions(range);
                if (size.Rows < 2 || columnIndex > size.Columns)
                    throw new ArgumentException("A filter requires a header plus data rows and an in-range columnIndex.");
                sheet = range.Worksheet;
                RejectTableFilterScope(ctx, sheet, range);
                filter = ResolveMatchingWorksheetFilter(sheet, range);
                if (filter is null && sheet.FilterMode)
                    throw new InvalidOperationException("An existing worksheet-wide advanced row filter has no inspectable scope. " +
                        "Use clear-filters with clear_advanced=true (CLI: --clear-advanced true) only when clearing it is authorized, " +
                        "before applying an ordinary filter.");
                NativeFilterHelpers.Apply(ctx.Book, range, columnIndex, filterOptions, token);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath, Action = "apply-filter" };
            }
            finally
            {
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public RangeFilterResult GetFilters(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        return batch.Execute((ctx, token) =>
        {
            Excel.Range? range = null;
            Excel.Worksheet? sheet = null;
            Excel.AutoFilter? filter = null;
            Excel.Range? header = null;
            try
            {
                token.ThrowIfCancellationRequested();
                range = ResolveFillRange(ctx, sheetName, rangeAddress);
                var size = GetContentDimensions(range);
                sheet = range.Worksheet;
                RejectTableFilterScope(ctx, sheet, range);
                filter = ResolveMatchingWorksheetFilter(sheet, range);
                header = range.Resize[1, size.Columns];
                object? values = header.Value2;
                var names = Enumerable.Range(0, size.Columns).Select(column =>
                    Convert.ToString(GetInspectionCell(values, 1, size.Columns, 0, column), CultureInfo.InvariantCulture) ?? "").ToList();
                var read = filter is null
                    ? names.Select((name, index) => new ColumnFilter { ColumnName = name, ColumnIndex = index + 1 }).ToList()
                    : NativeFilterHelpers.Read(filter, names, token);
                return new RangeFilterResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "get-filters",
                    RangeAddress = range.Address,
                    FilterEnabled = filter is not null,
                    WorksheetFilterMode = sheet.FilterMode,
                    AdvancedCriteriaAvailable = !sheet.FilterMode || filter is not null,
                    HasActiveFilters = read.Any(column => column.IsFiltered),
                    ColumnFilters = read
                };
            }
            finally
            {
                ComUtilities.Release(ref header);
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ClearFilters(IExcelBatch batch, string sheetName, string rangeAddress, bool clearAdvanced = false)
    {
        return batch.Execute((ctx, token) =>
        {
            Excel.Range? range = null;
            Excel.Worksheet? sheet = null;
            Excel.AutoFilter? filter = null;
            try
            {
                token.ThrowIfCancellationRequested();
                range = ResolveFillRange(ctx, sheetName, rangeAddress);
                sheet = range.Worksheet;
                RejectTableFilterScope(ctx, sheet, range);
                filter = ResolveMatchingWorksheetFilter(sheet, range);
                if (filter is not null && filter.FilterMode)
                    filter.ShowAllData();
                else if (filter is null && sheet.FilterMode)
                {
                    if (!clearAdvanced)
                        throw new InvalidOperationException("Excel does not expose the original advanced-filter scope. " +
                            "Use clear_advanced=true (CLI: --clear-advanced true) only when clearing worksheet-wide advanced row filtering is authorized.");
                    sheet.ShowAllData();
                }
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath, Action = "clear-filters" };
            }
            finally
            {
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref range);
            }
        });
    }

    private static Excel.AutoFilter? ResolveMatchingWorksheetFilter(Excel.Worksheet sheet, Excel.Range range)
    {
        Excel.AutoFilter? filter = null;
        Excel.Range? scope = null;
        try
        {
            if (!sheet.AutoFilterMode)
                return null;
            filter = sheet.AutoFilter;
            if (filter is null)
                throw new InvalidOperationException("Excel reports worksheet filtering without exposing its filter.");
            scope = filter.Range;
            if (!string.Equals(scope.Address, range.Address, StringComparison.Ordinal))
                throw new InvalidOperationException("The worksheet already has an AutoFilter on a different range. " +
                    "Use its exact range; this operation will not replace an unrelated filter.");
            var result = filter;
            filter = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref scope);
            ComUtilities.Release(ref filter);
        }
    }

    private static void RejectTableFilterScope(ExcelContext context, Excel.Worksheet sheet, Excel.Range range)
    {
        Excel.ListObjects? tables = null;
        try
        {
            tables = sheet.ListObjects;
            for (int index = 1; index <= tables.Count; index++)
            {
                Excel.ListObject? table = null;
                Excel.Range? tableRange = null;
                Excel.Range? overlap = null;
                try
                {
                    table = tables[index];
                    tableRange = table.Range;
                    overlap = context.App.Intersect(range, tableRange);
                    if (overlap is not null)
                        throw new ArgumentException("The filter scope intersects an Excel table; use table_column (CLI: tablecolumn).");
                }
                finally
                {
                    ComUtilities.Release(ref overlap);
                    ComUtilities.Release(ref tableRange);
                    ComUtilities.Release(ref table);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref tables);
        }
    }
}

/// <summary>Complete ordinary-range AutoFilter state for the explicitly selected scope.</summary>
public sealed class RangeFilterResult : OperationResult
{
    /// <summary>Absolute selected range, including its header.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Whether a native AutoFilter exists on this exact range.</summary>
    public bool FilterEnabled { get; set; }
    /// <summary>Native worksheet row-filter context, including advanced filters.</summary>
    public bool WorksheetFilterMode { get; set; }
    /// <summary>False when advanced row filtering exists but Excel does not expose its original criteria.</summary>
    public bool AdvancedCriteriaAvailable { get; set; }
    /// <summary>Whether any column filter is active.</summary>
    public bool HasActiveFilters { get; set; }
    /// <summary>Every column, including inactive columns, with native criteria coverage.</summary>
    public List<ColumnFilter> ColumnFilters { get; set; } = [];
}

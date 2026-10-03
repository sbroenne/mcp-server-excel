using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native advanced filtering destination mode.</summary>
public enum AdvancedFilterMode
{
    /// <summary>Hide nonmatching source rows.</summary>
    InPlace,
    /// <summary>Copy matching records to an explicit output.</summary>
    Copy
}

/// <summary>Native advanced filtering scope and conservative copy protection bounds.</summary>
public sealed class AdvancedFilterResult : OperationResult
{
    /// <summary>Absolute source rectangle, including headers.</summary>
    public string SourceRange { get; set; } = string.Empty;
    /// <summary>Absolute native criteria rectangle.</summary>
    public string CriteriaRange { get; set; } = string.Empty;
    /// <summary>Selected native mode.</summary>
    public AdvancedFilterMode Mode { get; set; }
    /// <summary>Maximum possible copied output checked before writing, not a claimed actual result extent.</summary>
    public string? CheckedDestinationRange { get; set; }
    /// <summary>Whether native unique filtering was requested.</summary>
    public bool UniqueOnly { get; set; }
}

public partial class RangeCommands
{
    /// <inheritdoc />
    public AdvancedFilterResult AdvancedFilter(IExcelBatch batch, string sheetName, string rangeAddress,
        string criteriaRange, AdvancedFilterMode mode, string? copyToRange = null, bool uniqueOnly = false,
        OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty)
    {
        if (!Enum.IsDefined(mode))
            throw new ArgumentOutOfRangeException(nameof(mode));
        ValidateOverwritePolicy(overwritePolicy);
        if ((mode == AdvancedFilterMode.Copy) != (copyToRange is not null))
            throw new ArgumentException("Copy requires copyToRange; InPlace must omit it.");
        return batch.Execute((ctx, token) =>
        {
            Excel.Range? source = null;
            Excel.Range? criteria = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? copyHeader = null;
            Excel.Worksheet? outputSheet = null;
            Excel.Range? destination = null;
            Excel.Range? sourceOverlap = null;
            Excel.Range? criteriaOverlap = null;
            Excel.Worksheet? criteriaSheet = null;
            Excel.AutoFilter? existingFilter = null;
            try
            {
                token.ThrowIfCancellationRequested();
                source = ResolveFillRange(ctx, sheetName, rangeAddress);
                criteria = ResolveFillRange(ctx, sheetName, criteriaRange);
                var size = GetContentDimensions(source);
                if (size.Rows < 2 || GetContentDimensions(criteria).Rows < 2)
                    throw new ArgumentException("Source and criteria require headers plus at least one data/criteria row.");
                sheet = source.Worksheet;
                RejectTableFilterScope(ctx, sheet, source);
                if (mode == AdvancedFilterMode.InPlace)
                {
                    existingFilter = ResolveMatchingWorksheetFilter(sheet, source);
                    if (existingFilter is null && sheet.FilterMode)
                        throw new InvalidOperationException("An existing worksheet-wide row filter has no inspectable source. " +
                            "Clear it explicitly before applying a new in-place advanced filter.");
                }
                if (mode == AdvancedFilterMode.Copy)
                {
                    copyHeader = ResolveFillRange(ctx, sheetName, copyToRange!);
                    var headerSize = GetContentDimensions(copyHeader);
                    if (headerSize.Rows != 1)
                        throw new ArgumentException("copyToRange must be one cell or one row of selected output headers.");
                    int columns = headerSize.Columns == 1 ? size.Columns : headerSize.Columns;
                    if ((long)copyHeader.Row + size.Rows - 1 > 1_048_576 ||
                        (long)copyHeader.Column + columns - 1 > 16_384)
                        throw new ArgumentException("The maximum advanced-filter output exceeds worksheet boundaries.");
                    outputSheet = copyHeader.Worksheet;
                    if (!string.Equals(outputSheet.Name, sheet.Name, StringComparison.Ordinal))
                        throw new ArgumentException("Advanced-filter copy output must be on the source worksheet.");
                    destination = copyHeader.Resize[size.Rows, columns];
                    if (RangeMergeDiscovery.GetMergeCellsState(destination.MergeCells) != false)
                        throw new ArgumentException("The maximum advanced-filter output intersects merged cells.");
                    sourceOverlap = ctx.App.Intersect(destination, source);
                    criteriaSheet = criteria.Worksheet;
                    if (string.Equals(criteriaSheet.Name, outputSheet.Name, StringComparison.Ordinal))
                        criteriaOverlap = ctx.App.Intersect(destination, criteria);
                    if (sourceOverlap is not null || criteriaOverlap is not null)
                        throw new ArgumentException("Advanced-filter copy output must not overlap source or criteria.");
                    EnsureDestinationWritable(ctx, destination, overwritePolicy, token);
                }
                token.ThrowIfCancellationRequested();
                source.AdvancedFilter(mode == AdvancedFilterMode.Copy
                        ? Excel.XlFilterAction.xlFilterCopy : Excel.XlFilterAction.xlFilterInPlace,
                    CriteriaRange: criteria, CopyToRange: copyHeader is null ? Type.Missing : copyHeader, Unique: uniqueOnly);
                return new AdvancedFilterResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "advanced-filter",
                    SourceRange = source.Address,
                    CriteriaRange = criteria.Address,
                    Mode = mode,
                    UniqueOnly = uniqueOnly,
                    CheckedDestinationRange = destination?.Address
                };
            }
            finally
            {
                ComUtilities.Release(ref criteriaOverlap);
                ComUtilities.Release(ref criteriaSheet);
                ComUtilities.Release(ref sourceOverlap);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref outputSheet);
                ComUtilities.Release(ref copyHeader);
                ComUtilities.Release(ref existingFilter);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref criteria);
                ComUtilities.Release(ref source);
            }
        });
    }
}

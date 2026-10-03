using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet page setup operations.
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult SetPageSetup(
        IExcelBatch batch,
        string sheetName,
        string? orientation = null,
        int? fitToPagesWide = null,
        int? fitToPagesTall = null,
        bool? centerHorizontally = null,
        bool? centerVertically = null,
        PageSetupOptions? pageSetupOptions = null)
    {
        var parsedOrientation = orientation is null ? (Excel.XlPageOrientation?)null : ParseOrientation(orientation);
        ValidatePageOptions(pageSetupOptions, fitToPagesWide, fitToPagesTall);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PageSetup? pageSetup = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                pageSetup = sheet.PageSetup;
                if (pageSetup == null)
                {
                    throw new InvalidOperationException($"Page setup could not be resolved for sheet '{sheetName}'.");
                }

                var printArea = NormalizePrintRange(ctx.Book, sheetName, pageSetupOptions?.PrintArea, null);
                var titleRows = NormalizePrintRange(ctx.Book, sheetName, pageSetupOptions?.PrintTitleRows, true);
                var titleColumns = NormalizePrintRange(ctx.Book, sheetName, pageSetupOptions?.PrintTitleColumns, false);
                if (parsedOrientation.HasValue)
                    pageSetup.Orientation = parsedOrientation.Value;

                if (fitToPagesWide.HasValue || fitToPagesTall.HasValue)
                {
                    pageSetup.Zoom = false;
                }

                if (fitToPagesWide.HasValue)
                {
                    pageSetup.FitToPagesWide = fitToPagesWide.Value == 0 ? false : (object)fitToPagesWide.Value;
                }

                if (fitToPagesTall.HasValue)
                {
                    pageSetup.FitToPagesTall = fitToPagesTall.Value == 0 ? false : (object)fitToPagesTall.Value;
                }

                if (centerHorizontally.HasValue)
                {
                    pageSetup.CenterHorizontally = centerHorizontally.Value;
                }

                if (centerVertically.HasValue)
                {
                    pageSetup.CenterVertically = centerVertically.Value;
                }

                if (pageSetupOptions is not null)
                    ApplyPageOptions(pageSetup, pageSetupOptions, printArea, titleRows, titleColumns);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref pageSetup);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SheetPageSetupResult GetPageSetup(IExcelBatch batch, string sheetName)
    {
        var result = new SheetPageSetupResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PageSetup? pageSetup = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                pageSetup = sheet.PageSetup;
                if (pageSetup == null)
                {
                    throw new InvalidOperationException($"Page setup could not be resolved for sheet '{sheetName}'.");
                }

                result.Orientation = GetOrientationName(pageSetup.Orientation);
                var fitToPagesEnabled = pageSetup.Zoom is bool zoom && !zoom;
                result.FitToPagesWide = fitToPagesEnabled ? GetFitToPagesValue(pageSetup.FitToPagesWide) : null;
                result.FitToPagesTall = fitToPagesEnabled ? GetFitToPagesValue(pageSetup.FitToPagesTall) : null;
                result.CenterHorizontally = pageSetup.CenterHorizontally;
                result.CenterVertically = pageSetup.CenterVertically;
                result.ZoomPercent = fitToPagesEnabled ? null : Convert.ToInt32(pageSetup.Zoom, CultureInfo.InvariantCulture);
                result.PrintArea = pageSetup.PrintArea ?? string.Empty;
                result.PrintTitleRows = pageSetup.PrintTitleRows ?? string.Empty;
                result.PrintTitleColumns = pageSetup.PrintTitleColumns ?? string.Empty;
                result.LeftMargin = pageSetup.LeftMargin;
                result.RightMargin = pageSetup.RightMargin;
                result.TopMargin = pageSetup.TopMargin;
                result.BottomMargin = pageSetup.BottomMargin;
                result.HeaderMargin = pageSetup.HeaderMargin;
                result.FooterMargin = pageSetup.FooterMargin;
                result.LeftHeader = pageSetup.LeftHeader ?? string.Empty;
                result.CenterHeader = pageSetup.CenterHeader ?? string.Empty;
                result.RightHeader = pageSetup.RightHeader ?? string.Empty;
                result.LeftFooter = pageSetup.LeftFooter ?? string.Empty;
                result.CenterFooter = pageSetup.CenterFooter ?? string.Empty;
                result.RightFooter = pageSetup.RightFooter ?? string.Empty;
                result.PaperSize = pageSetup.PaperSize.ToString();
                result.PageOrder = pageSetup.Order.ToString();
                result.PrintGridlines = pageSetup.PrintGridlines;
                result.PrintHeadings = pageSetup.PrintHeadings;
                result.BlackAndWhite = pageSetup.BlackAndWhite;
                result.Draft = pageSetup.Draft;
                result.FirstPageNumber = pageSetup.FirstPageNumber == (int)Excel.Constants.xlAutomatic ? 0 : pageSetup.FirstPageNumber;
                result.PrintComments = pageSetup.PrintComments.ToString();
                result.PrintErrors = pageSetup.PrintErrors.ToString();
                result.ScaleWithDocHeaderFooter = pageSetup.ScaleWithDocHeaderFooter;
                result.AlignMarginsHeaderFooter = pageSetup.AlignMarginsHeaderFooter;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref pageSetup);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static Excel.XlPageOrientation ParseOrientation(string orientation)
    {
        var normalized = (orientation ?? string.Empty).Trim().ToLowerInvariant();
        return normalized switch
        {
            "portrait" => Excel.XlPageOrientation.xlPortrait,
            "landscape" => Excel.XlPageOrientation.xlLandscape,
            _ => throw new ArgumentException($"Unsupported orientation '{orientation}'. Use 'portrait' or 'landscape'.", nameof(orientation))
        };
    }

    private static string GetOrientationName(Excel.XlPageOrientation orientation)
    {
        return orientation == Excel.XlPageOrientation.xlLandscape ? "landscape" : "portrait";
    }

    private static int? GetFitToPagesValue(object value)
    {
        return value is bool enabled && !enabled
            ? null
            : Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    private static void ValidatePageOptions(PageSetupOptions? options, int? wide, int? tall)
    {
        if (wide is < 0 || tall is < 0)
            throw new ArgumentOutOfRangeException(nameof(wide), "Fit page counts must be nonnegative; zero means unlimited.");
        if (options is null)
            return;
        if (options.ZoomPercent is < 10 or > 400)
            throw new ArgumentOutOfRangeException(nameof(options), "ZoomPercent must be between 10 and 400.");
        if (options.ZoomPercent.HasValue && (wide.HasValue || tall.HasValue))
            throw new ArgumentException("ZoomPercent and fit-to-page settings select conflicting scaling modes.");
        foreach (var margin in new[] { options.LeftMargin, options.RightMargin, options.TopMargin, options.BottomMargin, options.HeaderMargin, options.FooterMargin })
            if (margin.HasValue && (!double.IsFinite(margin.Value) || margin.Value < 0))
                throw new ArgumentException("Margins must be finite nonnegative point values.");
        foreach (var text in new[] { options.LeftHeader, options.CenterHeader, options.RightHeader, options.LeftFooter, options.CenterFooter, options.RightFooter })
            if (text is { Length: > 255 })
                throw new ArgumentException("Excel header/footer text must not exceed 255 characters, including formatting codes.");
        if (options.FirstPageNumber is < 0)
            throw new ArgumentException("FirstPageNumber must be nonnegative; zero selects automatic numbering.");
        _ = ParsePageEnum<Excel.XlPaperSize>(options.PaperSize);
        _ = ParsePageEnum<Excel.XlOrder>(options.PageOrder);
        _ = ParsePageEnum<Excel.XlPrintLocation>(options.PrintComments);
        _ = ParsePageEnum<Excel.XlPrintErrors>(options.PrintErrors);
    }

    private static T? ParsePageEnum<T>(string? value) where T : struct, Enum
    {
        if (value is null)
            return null;
        if (!Enum.GetNames<T>().Contains(value, StringComparer.OrdinalIgnoreCase) ||
            !Enum.TryParse<T>(value, true, out var parsed))
            throw new ArgumentException($"Unsupported native {typeof(T).Name} name '{value}'.");
        return parsed;
    }

    private static string? NormalizePrintRange(Excel.Workbook book, string sheetName, string? address, bool? rows)
    {
        if (address is null || address.Length == 0)
            return address;
        Excel.Range? range = null;
        Excel.Range? dimensions = null;
        Excel.Areas? areas = null;
        try
        {
            range = RangeHelpers.ResolveRange(book, sheetName, address, out _);
            if (rows.HasValue)
            {
                areas = range!.Areas;
                if (areas.Count != 1)
                    throw new ArgumentException("Print titles require one contiguous complete row or column range.");
                dimensions = rows.Value ? range.Columns : range.Rows;
                if (dimensions.Count != (rows.Value ? 16384 : 1048576))
                    throw new ArgumentException("Print title rows/columns must select complete rows/columns, not individual cells.");
            }
            return range!.Address;
        }
        finally
        {
            ComUtilities.Release(ref areas);
            ComUtilities.Release(ref dimensions);
            ComUtilities.Release(ref range);
        }
    }

    private static void ApplyPageOptions(Excel.PageSetup page, PageSetupOptions options, string? printArea, string? titleRows, string? titleColumns)
    {
        if (printArea is not null) page.PrintArea = printArea;
        if (titleRows is not null) page.PrintTitleRows = titleRows;
        if (titleColumns is not null) page.PrintTitleColumns = titleColumns;
        if (options.LeftMargin.HasValue) page.LeftMargin = options.LeftMargin.Value;
        if (options.RightMargin.HasValue) page.RightMargin = options.RightMargin.Value;
        if (options.TopMargin.HasValue) page.TopMargin = options.TopMargin.Value;
        if (options.BottomMargin.HasValue) page.BottomMargin = options.BottomMargin.Value;
        if (options.HeaderMargin.HasValue) page.HeaderMargin = options.HeaderMargin.Value;
        if (options.FooterMargin.HasValue) page.FooterMargin = options.FooterMargin.Value;
        if (options.LeftHeader is not null) page.LeftHeader = options.LeftHeader;
        if (options.CenterHeader is not null) page.CenterHeader = options.CenterHeader;
        if (options.RightHeader is not null) page.RightHeader = options.RightHeader;
        if (options.LeftFooter is not null) page.LeftFooter = options.LeftFooter;
        if (options.CenterFooter is not null) page.CenterFooter = options.CenterFooter;
        if (options.RightFooter is not null) page.RightFooter = options.RightFooter;
        if (options.PaperSize is not null) page.PaperSize = ParsePageEnum<Excel.XlPaperSize>(options.PaperSize)!.Value;
        if (options.PageOrder is not null) page.Order = ParsePageEnum<Excel.XlOrder>(options.PageOrder)!.Value;
        if (options.PrintGridlines.HasValue) page.PrintGridlines = options.PrintGridlines.Value;
        if (options.PrintHeadings.HasValue) page.PrintHeadings = options.PrintHeadings.Value;
        if (options.ZoomPercent.HasValue) page.Zoom = options.ZoomPercent.Value;
        if (options.BlackAndWhite.HasValue) page.BlackAndWhite = options.BlackAndWhite.Value;
        if (options.Draft.HasValue) page.Draft = options.Draft.Value;
        if (options.FirstPageNumber.HasValue) page.FirstPageNumber = options.FirstPageNumber.Value == 0 ? (int)Excel.Constants.xlAutomatic : options.FirstPageNumber.Value;
        if (options.PrintComments is not null) page.PrintComments = ParsePageEnum<Excel.XlPrintLocation>(options.PrintComments)!.Value;
        if (options.PrintErrors is not null) page.PrintErrors = ParsePageEnum<Excel.XlPrintErrors>(options.PrintErrors)!.Value;
        if (options.ScaleWithDocHeaderFooter.HasValue) page.ScaleWithDocHeaderFooter = options.ScaleWithDocHeaderFooter.Value;
        if (options.AlignMarginsHeaderFooter.HasValue) page.AlignMarginsHeaderFooter = options.AlignMarginsHeaderFooter.Value;
    }
}

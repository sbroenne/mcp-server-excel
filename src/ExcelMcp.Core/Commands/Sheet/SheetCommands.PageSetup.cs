using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
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
        string orientation,
        int? fitToPagesWide = null,
        int? fitToPagesTall = null,
        bool? centerHorizontally = null,
        bool? centerVertically = null)
    {
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

                pageSetup.Orientation = ParseOrientation(orientation);

                if (fitToPagesWide.HasValue || fitToPagesTall.HasValue)
                {
                    pageSetup.Zoom = false;
                }

                if (fitToPagesWide.HasValue)
                {
                    pageSetup.FitToPagesWide = fitToPagesWide.Value;
                }

                if (fitToPagesTall.HasValue)
                {
                    pageSetup.FitToPagesTall = fitToPagesTall.Value;
                }

                if (centerHorizontally.HasValue)
                {
                    pageSetup.CenterHorizontally = centerHorizontally.Value;
                }

                if (centerVertically.HasValue)
                {
                    pageSetup.CenterVertically = centerVertically.Value;
                }

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
}

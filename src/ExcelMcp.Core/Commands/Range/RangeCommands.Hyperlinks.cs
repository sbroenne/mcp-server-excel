using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range hyperlink operations (add, remove, list, get)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult AddHyperlink(
        IExcelBatch batch,
        string sheetName,
        string cellAddress,
        string? url = null,
        string? displayText = null,
        string? tooltip = null,
        string? subAddress = null)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "add-hyperlink" };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Hyperlinks? hyperlinks = null;
            Excel.Hyperlink? hyperlink = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                range = sheet.Range[cellAddress];
                hyperlinks = sheet.Hyperlinks;

                string address = NormalizeHyperlinkAddress(url, subAddress);

                hyperlink = hyperlinks.Add(
                    Anchor: range,
                    Address: address,
                    SubAddress: subAddress ?? Type.Missing,
                    ScreenTip: tooltip ?? Type.Missing,
                    TextToDisplay: displayText ?? Type.Missing
                );

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref hyperlink);
                ComUtilities.Release(ref hyperlinks);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult UpdateHyperlink(
        IExcelBatch batch,
        string sheetName,
        string cellAddress,
        string? url = null,
        string? displayText = null,
        string? tooltip = null,
        string? subAddress = null)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "update-hyperlink" };

        return batch.Execute((ctx, ct) =>
        {
            if (url == null && displayText == null && tooltip == null && subAddress == null)
            {
                throw new ArgumentException("Provide at least one hyperlink property to update.");
            }

            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Hyperlinks? hyperlinks = null;
            Excel.Hyperlink? hyperlink = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                range = sheet.Range[cellAddress];
                hyperlinks = range.Hyperlinks;
                if (hyperlinks.Count == 0)
                {
                    throw new InvalidOperationException($"Cell '{sheetName}!{cellAddress}' does not contain a hyperlink.");
                }

                hyperlink = hyperlinks[1];
                string existingAddress = hyperlink.Address ?? string.Empty;
                string effectiveSubAddress = subAddress ?? hyperlink.SubAddress ?? string.Empty;
                string effectiveAddress = url != null
                    ? NormalizeHyperlinkAddress(url, effectiveSubAddress)
                    : existingAddress;
                string effectiveDisplayText = displayText ?? hyperlink.TextToDisplay ?? string.Empty;
                string? effectiveTooltip = tooltip ?? hyperlink.ScreenTip;

                if (string.IsNullOrWhiteSpace(effectiveAddress) && string.IsNullOrWhiteSpace(effectiveSubAddress))
                {
                    throw new ArgumentException("A hyperlink must retain either an external address or an internal subAddress.");
                }

                if (url != null || subAddress != null)
                {
                    hyperlink.Delete();
                    ComUtilities.Release(ref hyperlink);
                    hyperlink = hyperlinks.Add(
                        Anchor: range,
                        Address: effectiveAddress,
                        SubAddress: string.IsNullOrEmpty(effectiveSubAddress) ? Type.Missing : effectiveSubAddress,
                        ScreenTip: effectiveTooltip ?? Type.Missing,
                        TextToDisplay: effectiveDisplayText);
                }
                else
                {
                    if (displayText != null)
                    {
                        hyperlink.TextToDisplay = displayText;
                    }

                    if (tooltip != null)
                    {
                        hyperlink.ScreenTip = tooltip;
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref hyperlink);
                ComUtilities.Release(ref hyperlinks);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult RemoveHyperlink(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "remove-hyperlink" };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Hyperlinks? hyperlinks = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                hyperlinks = range.Hyperlinks;
                int count = hyperlinks.Count;

                // Delete all hyperlinks in the range
                for (int i = count; i >= 1; i--)
                {
                    Excel.Hyperlink? hl = null;
                    try
                    {
                        hl = hyperlinks[i];
                        hl.Delete();
                    }
                    finally
                    {
                        ComUtilities.Release(ref hl);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref hyperlinks);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public RangeHyperlinkResult ListHyperlinks(IExcelBatch batch, string sheetName)
    {
        var result = new RangeHyperlinkResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Hyperlinks? hyperlinks = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                hyperlinks = sheet.Hyperlinks;
                int count = hyperlinks.Count;

                for (int i = 1; i <= count; i++)
                {
                    Excel.Hyperlink? hyperlink = null;
                    Excel.Range? range = null;
                    try
                    {
                        hyperlink = hyperlinks[i];
                        range = hyperlink.Range;

                        result.Hyperlinks.Add(new HyperlinkInfo
                        {
                            CellAddress = range.Address[false, false],
                            Address = hyperlink.Address ?? string.Empty,
                            SubAddress = hyperlink.SubAddress,
                            DisplayText = hyperlink.TextToDisplay ?? string.Empty,
                            ScreenTip = hyperlink.ScreenTip,
                            IsInternal = string.IsNullOrEmpty(hyperlink.Address)
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref range);
                        ComUtilities.Release(ref hyperlink);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref hyperlinks);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public RangeHyperlinkResult GetHyperlink(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new RangeHyperlinkResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = cellAddress
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Hyperlinks? hyperlinks = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                range = sheet.Range[cellAddress];
                hyperlinks = range.Hyperlinks;

                int count = hyperlinks.Count;
                if (count > 0)
                {
                    Excel.Hyperlink? hyperlink = null;
                    try
                    {
                        hyperlink = hyperlinks[1]; // Get first hyperlink in cell

                        result.Hyperlinks.Add(new HyperlinkInfo
                        {
                            CellAddress = cellAddress,
                            Address = hyperlink.Address ?? string.Empty,
                            SubAddress = hyperlink.SubAddress,
                            DisplayText = hyperlink.TextToDisplay ?? string.Empty,
                            ScreenTip = hyperlink.ScreenTip,
                            IsInternal = string.IsNullOrEmpty(hyperlink.Address)
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref hyperlink);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref hyperlinks);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static string NormalizeHyperlinkAddress(string? url, string? subAddress)
    {
        if (string.IsNullOrWhiteSpace(url))
        {
            if (string.IsNullOrWhiteSpace(subAddress))
            {
                throw new ArgumentException("Provide url for an external hyperlink or subAddress for an internal hyperlink.");
            }

            return string.Empty;
        }

        return url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
               url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
               url.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
               url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
            ? url
            : Path.GetFullPath(url);
    }
}

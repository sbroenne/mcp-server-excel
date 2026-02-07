using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range hyperlink operations (add, remove, list, get)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult AddHyperlink(IExcelBatch batch, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "add-hyperlink" };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? hyperlinks = null;
            dynamic? hyperlink = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                range = sheet.Range[cellAddress];
                hyperlinks = sheet.Hyperlinks;

                // Resolve URL - full path for file links
                string address = url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                               url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                               url.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
                               url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
                    ? url
                    : Path.GetFullPath(url);

                hyperlink = hyperlinks.Add(
                    Anchor: range,
                    Address: address,
                    SubAddress: Type.Missing,
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
    public OperationResult RemoveHyperlink(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "remove-hyperlink" };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            dynamic? hyperlinks = null;
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
                    dynamic? hl = null;
                    try
                    {
                        hl = hyperlinks.Item(i);
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
            dynamic? sheet = null;
            dynamic? hyperlinks = null;
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
                    dynamic? hyperlink = null;
                    dynamic? range = null;
                    try
                    {
                        hyperlink = hyperlinks.Item(i);
                        range = hyperlink.Range;

                        result.Hyperlinks.Add(new HyperlinkInfo
                        {
                            CellAddress = range.Address[false, false],
                            Address = hyperlink.Address ?? string.Empty,
                            DisplayText = hyperlink.TextToDisplay ?? string.Empty,
                            ScreenTip = hyperlink.ScreenTip
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
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? hyperlinks = null;
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
                    dynamic? hyperlink = null;
                    try
                    {
                        hyperlink = hyperlinks.Item(1); // Get first hyperlink in cell

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

}




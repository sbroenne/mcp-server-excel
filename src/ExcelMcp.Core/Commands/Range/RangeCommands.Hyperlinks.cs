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
    public async Task<OperationResult> AddHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "add-hyperlink" };

        return await batch.Execute((ctx, ct) =>
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
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                range = sheet.Range[cellAddress];
                hyperlinks = sheet.Hyperlinks;

                // Resolve URL - full path for file links
                string address = url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                               url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                               url.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
                               url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
                    ? url
                    : System.IO.Path.GetFullPath(url);

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
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
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
    public async Task<OperationResult> RemoveHyperlinkAsync(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "remove-hyperlink" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? range = null;
            dynamic? hyperlinks = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
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
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
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
    public async Task<RangeHyperlinkResult> ListHyperlinksAsync(IExcelBatch batch, string sheetName)
    {
        var result = new RangeHyperlinkResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? hyperlinks = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
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
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
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
    public async Task<RangeHyperlinkResult> GetHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new RangeHyperlinkResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = cellAddress
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? hyperlinks = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
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
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
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

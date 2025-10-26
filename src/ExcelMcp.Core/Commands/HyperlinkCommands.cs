using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using Sbroenne.ExcelMcp.Core.Session;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Implementation of hyperlink-related commands using Excel COM interop.
/// </summary>
public class HyperlinkCommands : IHyperlinkCommands
{
    /// <summary>
    /// Adds a hyperlink to a cell or range in an Excel worksheet.
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell or range address (e.g., "A1" or "A1:B2")</param>
    /// <param name="url">The URL or file path to link to</param>
    /// <param name="displayText">Optional display text for the hyperlink</param>
    /// <param name="tooltip">Optional tooltip text</param>
    /// <returns>Operation result with success status and details</returns>
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> AddHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "hyperlink" };

        try
        {
return await batch.ExecuteAsync(async (ctx, ct) =>
            {
                dynamic? sheet = null;
                dynamic? range = null;
                dynamic? hyperlinks = null;
                dynamic? hyperlink = null;

                try
                {
                    sheet = ctx.Book.Worksheets.Item(sheetName);
                    range = sheet.Range[cellAddress];
                    hyperlinks = sheet.Hyperlinks;

                    // Add hyperlink
                    string anchor = range;
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
                    result.FilePath = batch.WorkbookPath;
                    result.WorkflowHint = $"Hyperlink added to {cellAddress} in sheet '{sheetName}'";
                    result.SuggestedNextActions = new List<string>
                    {
                        "List all hyperlinks in the sheet",
                        "Add more hyperlinks to other cells",
                        "Test the hyperlink by opening the file"
                    };

                    result.Success = true; return result;
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
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to add hyperlink: {ex.Message}";
            result.FilePath = batch.WorkbookPath;
        }

        return result;
    }

    /// <summary>
    /// Removes all hyperlinks from a cell or range in an Excel worksheet.
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell or range address (e.g., "A1" or "A1:B2")</param>
    /// <returns>Operation result with success status and details</returns>
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> RemoveHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "hyperlink" };

        try
        {
return await batch.ExecuteAsync(async (ctx, ct) =>
            {
                dynamic? sheet = null;
                dynamic? range = null;
                dynamic? hyperlinks = null;

                try
                {
                    sheet = ctx.Book.Worksheets.Item(sheetName);
                    range = sheet.Range[cellAddress];
                    hyperlinks = range.Hyperlinks;

                    int count = hyperlinks.Count;
                    if (count > 0)
                    {
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
                        result.FilePath = batch.WorkbookPath;
                        result.WorkflowHint = $"Removed {count} hyperlink(s) from {cellAddress} in sheet '{sheetName}'";
                    }
                    else
                    {
                        result.Success = true;
                        result.FilePath = batch.WorkbookPath;
                        result.WorkflowHint = $"No hyperlinks found at {cellAddress} in sheet '{sheetName}'";
                    }

                    result.Success = true; return result;
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
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to remove hyperlink: {ex.Message}";
            result.FilePath = batch.WorkbookPath;
        }

        return result;
    }

    /// <summary>
    /// Lists all hyperlinks in a worksheet with their addresses and target URLs.
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <returns>Hyperlink list result with details of all hyperlinks in the sheet</returns>
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<HyperlinkListResult> ListHyperlinksAsync(IExcelBatch batch, string sheetName)
    {
        var result = new HyperlinkListResult 
        { 
            FilePath = batch.WorkbookPath,
            SheetName = sheetName
        };

        try
        {
return await batch.ExecuteAsync(async (ctx, ct) =>
            {
                dynamic? sheet = null;
                dynamic? hyperlinks = null;

                try
                {
                    sheet = ctx.Book.Worksheets.Item(sheetName);
                    hyperlinks = sheet.Hyperlinks;

                    int count = hyperlinks.Count;
                    result.Count = count;

                    for (int i = 1; i <= count; i++)
                    {
                        dynamic? hyperlink = null;
                        dynamic? range = null;

                        try
                        {
                            hyperlink = hyperlinks.Item(i);
                            range = hyperlink.Range;

                            var info = new HyperlinkInfo
                            {
                                CellAddress = range.Address[false, false],
                                Address = hyperlink.Address ?? string.Empty,
                                SubAddress = hyperlink.SubAddress,
                                DisplayText = hyperlink.TextToDisplay ?? string.Empty,
                                ScreenTip = hyperlink.ScreenTip,
                                IsInternal = string.IsNullOrEmpty(hyperlink.Address)
                            };

                            result.Hyperlinks.Add(info);
                        }
                        finally
                        {
                            ComUtilities.Release(ref range);
                            ComUtilities.Release(ref hyperlink);
                        }
                    }

                    result.Success = true;
                    result.WorkflowHint = count > 0
                        ? $"Found {count} hyperlink(s) in sheet '{sheetName}'"
                        : $"No hyperlinks found in sheet '{sheetName}'";

                    if (count > 0)
                    {
                        result.SuggestedNextActions = new List<string>
                        {
                            "Get detailed info for a specific hyperlink",
                            "Remove unwanted hyperlinks",
                            "Add more hyperlinks to other cells"
                        };
                    }

                    result.Success = true; return result;
                }
                finally
                {
                    ComUtilities.Release(ref hyperlinks);
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to list hyperlinks: {ex.Message}";
        }

        return result;
    }

    /// <summary>
    /// Gets hyperlink information from a specific cell.
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address (e.g., "A1")</param>
    /// <returns>Hyperlink info result with details of the hyperlink at the specified cell</returns>
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<HyperlinkInfoResult> GetHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new HyperlinkInfoResult 
        { 
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            CellAddress = cellAddress
        };

        try
        {
return await batch.ExecuteAsync(async (ctx, ct) =>
            {
                dynamic? sheet = null;
                dynamic? range = null;
                dynamic? hyperlinks = null;

                try
                {
                    sheet = ctx.Book.Worksheets.Item(sheetName);
                    range = sheet.Range[cellAddress];
                    hyperlinks = range.Hyperlinks;

                    int count = hyperlinks.Count;
                    result.HasHyperlink = count > 0;

                    if (count > 0)
                    {
                        dynamic? hyperlink = null;
                        try
                        {
                            hyperlink = hyperlinks.Item(1); // Get first hyperlink in cell

                            result.Hyperlink = new HyperlinkInfo
                            {
                                CellAddress = cellAddress,
                                Address = hyperlink.Address ?? string.Empty,
                                SubAddress = hyperlink.SubAddress,
                                DisplayText = hyperlink.TextToDisplay ?? string.Empty,
                                ScreenTip = hyperlink.ScreenTip,
                                IsInternal = string.IsNullOrEmpty(hyperlink.Address)
                            };

                            result.Success = true;
                            result.WorkflowHint = $"Hyperlink found at {cellAddress}: {result.Hyperlink.DisplayText}";
                        }
                        finally
                        {
                            ComUtilities.Release(ref hyperlink);
                        }
                    }
                    else
                    {
                        result.Success = true;
                        result.WorkflowHint = $"No hyperlink found at {cellAddress} in sheet '{sheetName}'";
                    }

                    result.Success = true; return result;
                }
                finally
                {
                    ComUtilities.Release(ref hyperlinks);
                    ComUtilities.Release(ref range);
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to get hyperlink info: {ex.Message}";
        }

        return result;
    }
}




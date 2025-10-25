using ExcelMcp.Core.Models;
using ExcelMcp.Core.Utils;
using System.Runtime.InteropServices;

namespace ExcelMcp.Core.Commands;

/// <summary>
/// Implementation of hyperlink-related commands using Excel COM interop.
/// </summary>
public class HyperlinkCommands : IHyperlinkCommands
{
    public Result AddHyperlink(string excelPath, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null)
    {
        var result = new Result();

        try
        {
            PathValidator.ValidateFilePath(excelPath, allowCreate: false);

            return ExcelHelper.WithExcel(excelPath, save: true, (excel, workbook) =>
            {
                dynamic? sheet = null;
                dynamic? range = null;
                dynamic? hyperlinks = null;
                dynamic? hyperlink = null;

                try
                {
                    sheet = workbook.Worksheets.Item(sheetName);
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
                    result.FilePath = excelPath;
                    result.WorkflowHint = $"Hyperlink added to {cellAddress} in sheet '{sheetName}'";
                    result.SuggestedNextActions = new List<string>
                    {
                        "List all hyperlinks in the sheet",
                        "Add more hyperlinks to other cells",
                        "Test the hyperlink by opening the file"
                    };

                    return 0;
                }
                finally
                {
                    ExcelHelper.ReleaseComObject(ref hyperlink);
                    ExcelHelper.ReleaseComObject(ref hyperlinks);
                    ExcelHelper.ReleaseComObject(ref range);
                    ExcelHelper.ReleaseComObject(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to add hyperlink: {ex.Message}";
            result.FilePath = excelPath;
        }

        return result;
    }

    public Result RemoveHyperlink(string excelPath, string sheetName, string cellAddress)
    {
        var result = new Result();

        try
        {
            PathValidator.ValidateFilePath(excelPath, allowCreate: false);

            return ExcelHelper.WithExcel(excelPath, save: true, (excel, workbook) =>
            {
                dynamic? sheet = null;
                dynamic? range = null;
                dynamic? hyperlinks = null;

                try
                {
                    sheet = workbook.Worksheets.Item(sheetName);
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
                                ExcelHelper.ReleaseComObject(ref hl);
                            }
                        }

                        result.Success = true;
                        result.FilePath = excelPath;
                        result.WorkflowHint = $"Removed {count} hyperlink(s) from {cellAddress} in sheet '{sheetName}'";
                    }
                    else
                    {
                        result.Success = true;
                        result.FilePath = excelPath;
                        result.WorkflowHint = $"No hyperlinks found at {cellAddress} in sheet '{sheetName}'";
                    }

                    return 0;
                }
                finally
                {
                    ExcelHelper.ReleaseComObject(ref hyperlinks);
                    ExcelHelper.ReleaseComObject(ref range);
                    ExcelHelper.ReleaseComObject(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to remove hyperlink: {ex.Message}";
            result.FilePath = excelPath;
        }

        return result;
    }

    public HyperlinkListResult ListHyperlinks(string excelPath, string sheetName)
    {
        var result = new HyperlinkListResult
        {
            SheetName = sheetName,
            FilePath = excelPath
        };

        try
        {
            PathValidator.ValidateFilePath(excelPath, allowCreate: false);

            ExcelHelper.WithExcel(excelPath, save: false, (excel, workbook) =>
            {
                dynamic? sheet = null;
                dynamic? hyperlinks = null;

                try
                {
                    sheet = workbook.Worksheets.Item(sheetName);
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
                            ExcelHelper.ReleaseComObject(ref range);
                            ExcelHelper.ReleaseComObject(ref hyperlink);
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

                    return 0;
                }
                finally
                {
                    ExcelHelper.ReleaseComObject(ref hyperlinks);
                    ExcelHelper.ReleaseComObject(ref sheet);
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

    public HyperlinkInfoResult GetHyperlink(string excelPath, string sheetName, string cellAddress)
    {
        var result = new HyperlinkInfoResult
        {
            SheetName = sheetName,
            CellAddress = cellAddress,
            FilePath = excelPath
        };

        try
        {
            PathValidator.ValidateFilePath(excelPath, allowCreate: false);

            ExcelHelper.WithExcel(excelPath, save: false, (excel, workbook) =>
            {
                dynamic? sheet = null;
                dynamic? range = null;
                dynamic? hyperlinks = null;

                try
                {
                    sheet = workbook.Worksheets.Item(sheetName);
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
                            ExcelHelper.ReleaseComObject(ref hyperlink);
                        }
                    }
                    else
                    {
                        result.Success = true;
                        result.WorkflowHint = $"No hyperlink found at {cellAddress} in sheet '{sheetName}'";
                    }

                    return 0;
                }
                finally
                {
                    ExcelHelper.ReleaseComObject(ref hyperlinks);
                    ExcelHelper.ReleaseComObject(ref range);
                    ExcelHelper.ReleaseComObject(ref sheet);
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

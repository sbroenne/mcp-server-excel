using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// FilePath-based worksheet operations using FileHandleManager
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public async Task<WorksheetListResult> ListAsync(string filePath)
    {
        var result = new WorksheetListResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheets = null;
                try
                {
                    sheets = handle.Workbook.Worksheets;
                    for (int i = 1; i <= sheets.Count; i++)
                    {
                        dynamic? sheet = null;
                        try
                        {
                            sheet = sheets.Item(i);
                            result.Worksheets.Add(new WorksheetInfo { Name = sheet.Name, Index = i });
                        }
                        finally
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheets);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(string filePath, string sheetName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "create-sheet" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheets = null;
                dynamic? newSheet = null;
                try
                {
                    sheets = handle.Workbook.Worksheets;
                    newSheet = sheets.Add();
                    newSheet.Name = sheetName;
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref newSheet);
                    ComUtilities.Release(ref sheets);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> RenameAsync(string filePath, string oldName, string newName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "rename-sheet" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheet = null;
                try
                {
                    sheet = handle.Workbook.Worksheets[oldName];
                    sheet.Name = newName;
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> CopyAsync(string filePath, string sourceName, string targetName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "copy-sheet" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sourceSheet = null;
                dynamic? copiedSheet = null;
                try
                {
                    sourceSheet = handle.Workbook.Worksheets[sourceName];
                    sourceSheet.Copy(After: sourceSheet);
                    copiedSheet = handle.Workbook.Worksheets[sourceSheet.Index + 1];
                    copiedSheet.Name = targetName;
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref copiedSheet);
                    ComUtilities.Release(ref sourceSheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string sheetName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "delete-sheet" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheet = null;
                try
                {
                    sheet = handle.Workbook.Worksheets[sheetName];
                    sheet.Delete();
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetTabColorAsync(string filePath, string sheetName, int red, int green, int blue)
    {
        var result = new OperationResult { FilePath = filePath, Action = "set-tab-color" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheet = null;
                try
                {
                    sheet = handle.Workbook.Worksheets[sheetName];
                    // Convert RGB to BGR format (Excel uses BGR internally)
                    int bgrColor = (blue << 16) | (green << 8) | red;
                    sheet.Tab.Color = bgrColor;
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<TabColorResult> GetTabColorAsync(string filePath, string sheetName)
    {
        var result = new TabColorResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheet = null;
                try
                {
                    sheet = handle.Workbook.Worksheets[sheetName];
                    object colorValue = sheet.Tab.Color;

                    if (colorValue == null || colorValue is DBNull)
                    {
                        result.HasColor = false;
                    }
                    else
                    {
                        int bgrColor = Convert.ToInt32(colorValue, CultureInfo.InvariantCulture);
                        // Convert BGR to RGB
                        int red = bgrColor & 0xFF;
                        int green = (bgrColor >> 8) & 0xFF;
                        int blue = (bgrColor >> 16) & 0xFF;

                        result.HasColor = true;
                        result.Red = red;
                        result.Green = green;
                        result.Blue = blue;
                        result.HexColor = $"#{red:X2}{green:X2}{blue:X2}";
                    }

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> ClearTabColorAsync(string filePath, string sheetName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "clear-tab-color" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheet = null;
                try
                {
                    sheet = handle.Workbook.Worksheets[sheetName];
                    sheet.Tab.Color = false; // false clears the color
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetVisibilityAsync(string filePath, string sheetName, SheetVisibility visibility)
    {
        var result = new OperationResult { FilePath = filePath, Action = "set-visibility" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheet = null;
                try
                {
                    sheet = handle.Workbook.Worksheets[sheetName];
                    sheet.Visible = (int)visibility;
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<SheetVisibilityResult> GetVisibilityAsync(string filePath, string sheetName)
    {
        var result = new SheetVisibilityResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? sheet = null;
                try
                {
                    sheet = handle.Workbook.Worksheets[sheetName];
                    int visibilityValue = sheet.Visible;
                    result.Visibility = (SheetVisibility)visibilityValue;
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <inheritdoc />
    public Task<OperationResult> ShowAsync(string filePath, string sheetName)
    {
        return SetVisibilityAsync(filePath, sheetName, SheetVisibility.Visible);
    }

    /// <inheritdoc />
    public Task<OperationResult> HideAsync(string filePath, string sheetName)
    {
        return SetVisibilityAsync(filePath, sheetName, SheetVisibility.Hidden);
    }

    /// <inheritdoc />
    public Task<OperationResult> VeryHideAsync(string filePath, string sheetName)
    {
        return SetVisibilityAsync(filePath, sheetName, SheetVisibility.VeryHidden);
    }
}

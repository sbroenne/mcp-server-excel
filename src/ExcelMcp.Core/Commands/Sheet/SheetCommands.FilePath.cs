using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet operations - FilePath-based API implementations
/// </summary>
public partial class SheetCommands
{
    #region Lifecycle Operations

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
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheets);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref newSheet);
                    ComUtilities.Release(ref sheets);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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
                    sheet = ComUtilities.FindSheet(handle.Workbook, oldName);
                    if (sheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Sheet '{oldName}' not found";
                        return;
                    }

                    sheet.Name = newName;
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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
                    sourceSheet = ComUtilities.FindSheet(handle.Workbook, sourceName);
                    if (sourceSheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Source sheet '{sourceName}' not found";
                        return;
                    }

                    sourceSheet.Copy(After: sourceSheet);
                    dynamic? sheets = null;
                    try
                    {
                        sheets = handle.Workbook.Worksheets;
                        copiedSheet = sheets.Item(sourceSheet.Index + 1);
                        copiedSheet.Name = targetName;
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheets);
                    }

                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref copiedSheet);
                    ComUtilities.Release(ref sourceSheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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
                    sheet = ComUtilities.FindSheet(handle.Workbook, sheetName);
                    if (sheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Sheet '{sheetName}' not found";
                        return;
                    }

                    sheet.Delete();
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    #endregion

    #region Tab Color Operations

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
                    sheet = ComUtilities.FindSheet(handle.Workbook, sheetName);
                    if (sheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Sheet '{sheetName}' not found";
                        return;
                    }

                    dynamic? tab = null;
                    try
                    {
                        tab = sheet.Tab;
                        int bgrColor = (blue << 16) | (green << 8) | red;
                        tab.Color = bgrColor;
                        result.Success = true;
                    }
                    finally
                    {
                        ComUtilities.Release(ref tab);
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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
                    sheet = ComUtilities.FindSheet(handle.Workbook, sheetName);
                    if (sheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Sheet '{sheetName}' not found";
                        return;
                    }

                    dynamic? tab = null;
                    try
                    {
                        tab = sheet.Tab;
                        object colorObj = tab.Color;

                        if (colorObj == null || Convert.ToInt32(colorObj, System.Globalization.CultureInfo.InvariantCulture) == 0)
                        {
                            result.HasColor = false;
                        }
                        else
                        {
                            int bgrColor = Convert.ToInt32(colorObj, System.Globalization.CultureInfo.InvariantCulture);
                            result.Blue = (bgrColor >> 16) & 0xFF;
                            result.Green = (bgrColor >> 8) & 0xFF;
                            result.Red = bgrColor & 0xFF;
                            result.HexColor = $"#{result.Red:X2}{result.Green:X2}{result.Blue:X2}";
                            result.HasColor = true;
                        }

                        result.Success = true;
                    }
                    finally
                    {
                        ComUtilities.Release(ref tab);
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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
                    sheet = ComUtilities.FindSheet(handle.Workbook, sheetName);
                    if (sheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Sheet '{sheetName}' not found";
                        return;
                    }

                    dynamic? tab = null;
                    try
                    {
                        tab = sheet.Tab;
                        tab.Color = System.Reflection.Missing.Value;
                        result.Success = true;
                    }
                    finally
                    {
                        ComUtilities.Release(ref tab);
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    #endregion

    #region Visibility Operations

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
                    sheet = ComUtilities.FindSheet(handle.Workbook, sheetName);
                    if (sheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Sheet '{sheetName}' not found";
                        return;
                    }

                    sheet.Visible = (int)visibility;
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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
                    sheet = ComUtilities.FindSheet(handle.Workbook, sheetName);
                    if (sheet == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Sheet '{sheetName}' not found";
                        return;
                    }

                    int visibleValue = Convert.ToInt32(sheet.Visible, System.Globalization.CultureInfo.InvariantCulture);
                    result.Visibility = (SheetVisibility)visibleValue;
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
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

    #endregion
}

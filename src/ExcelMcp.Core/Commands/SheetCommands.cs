using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle and appearance management commands.
/// - Lifecycle: create, rename, copy, delete worksheets
/// - Appearance: tab colors, visibility levels
/// Data operations (read, write, clear) moved to Range.IRangeCommands for unified range API.
/// All operations use batching for performance.
/// </summary>
public class SheetCommands : ISheetCommands
{
    /// <inheritdoc />
    public async Task<WorksheetListResult> ListAsync(IExcelBatch batch)
    {
        var result = new WorksheetListResult { FilePath = batch.WorkbookPath };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
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
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "create-sheet" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? newSheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                newSheet = sheets.Add();
                newSheet.Name = sheetName;
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
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> RenameAsync(IExcelBatch batch, string oldName, string newName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "rename-sheet" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, oldName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{oldName}' not found";
                    return result;
                }
                sheet.Name = newName;
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
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceName, string targetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "copy-sheet" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sourceSheet = null;
            dynamic? sheets = null;
            dynamic? lastSheet = null;
            dynamic? copiedSheet = null;
            try
            {
                sourceSheet = ComUtilities.FindSheet(ctx.Book, sourceName);
                if (sourceSheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sourceName}' not found";
                    return result;
                }
                sheets = ctx.Book.Worksheets;
                lastSheet = sheets.Item(sheets.Count);
                sourceSheet.Copy(After: lastSheet);
                copiedSheet = sheets.Item(sheets.Count);
                copiedSheet.Name = targetName;
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
                ComUtilities.Release(ref copiedSheet);
                ComUtilities.Release(ref lastSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref sourceSheet);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-sheet" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                sheet.Delete();
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
                ComUtilities.Release(ref sheet);
            }
        });
    }
    
    // === TAB COLOR OPERATIONS ===
    
    /// <inheritdoc />
    public async Task<OperationResult> SetTabColorAsync(IExcelBatch batch, string sheetName, int red, int green, int blue)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-tab-color" };
        
        // Validate RGB values
        if (red < 0 || red > 255 || green < 0 || green > 255 || blue < 0 || blue > 255)
        {
            result.Success = false;
            result.ErrorMessage = "RGB values must be between 0 and 255";
            return result;
        }

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? tab = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                
                // Convert RGB to BGR format (Excel's color format)
                // BGR = (Blue << 16) | (Green << 8) | Red
                int bgrColor = (blue << 16) | (green << 8) | red;
                
                tab = sheet.Tab;
                tab.Color = bgrColor;
                
                result.Success = true;
                result.SuggestedNextActions = 
                [
                    $"Tab color set to RGB({red}, {green}, {blue})",
                    "Use 'get-tab-color' to verify the color",
                    "Use 'clear-tab-color' to remove the color"
                ];
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
                ComUtilities.Release(ref tab);
                ComUtilities.Release(ref sheet);
            }
        });
    }
    
    /// <inheritdoc />
    public async Task<TabColorResult> GetTabColorAsync(IExcelBatch batch, string sheetName)
    {
        var result = new TabColorResult { FilePath = batch.WorkbookPath };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? tab = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                
                tab = sheet.Tab;
                dynamic colorValue = tab.Color;
                
                // Excel's ColorIndex.xlColorIndexAutomatic = -4105
                // When no custom color is set, Excel might return various values
                // Check ColorIndex property instead for more reliable detection
                dynamic colorIndex = tab.ColorIndex;
                
                // xlColorIndexNone = -4142, xlColorIndexAutomatic = -4105
                // If ColorIndex is negative or color value indicates no custom color
                if (colorIndex is int idx && (idx == -4142 || idx == -4105 || idx < 0))
                {
                    result.Success = true;
                    result.HasColor = false;
                    return result;
                }
                
                // Also check if color value itself indicates no custom color
                if (colorValue == null || colorValue == 0)
                {
                    result.Success = true;
                    result.HasColor = false;
                    return result;
                }
                
                // Convert BGR to RGB
                int bgrColor = Convert.ToInt32(colorValue);
                int red = bgrColor & 0xFF;
                int green = (bgrColor >> 8) & 0xFF;
                int blue = (bgrColor >> 16) & 0xFF;
                
                result.Success = true;
                result.HasColor = true;
                result.Red = red;
                result.Green = green;
                result.Blue = blue;
                result.HexColor = $"#{red:X2}{green:X2}{blue:X2}";
                
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
                ComUtilities.Release(ref tab);
                ComUtilities.Release(ref sheet);
            }
        });
    }
    
    /// <inheritdoc />
    public async Task<OperationResult> ClearTabColorAsync(IExcelBatch batch, string sheetName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "clear-tab-color" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? tab = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                
                tab = sheet.Tab;
                // Set ColorIndex to xlColorIndexNone (-4142) to clear color
                tab.ColorIndex = -4142; // xlColorIndexNone
                
                result.Success = true;
                result.SuggestedNextActions = 
                [
                    "Tab color cleared (reset to default)",
                    "Use 'set-tab-color' to apply a new color"
                ];
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
                ComUtilities.Release(ref tab);
                ComUtilities.Release(ref sheet);
            }
        });
    }
    
    // === VISIBILITY OPERATIONS ===
    
    /// <inheritdoc />
    public async Task<OperationResult> SetVisibilityAsync(IExcelBatch batch, string sheetName, SheetVisibility visibility)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-visibility" };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                
                // Set visibility using the enum value (maps to XlSheetVisibility)
                sheet.Visible = (int)visibility;
                
                result.Success = true;
                result.SuggestedNextActions = visibility switch
                {
                    SheetVisibility.Visible => 
                    [
                        $"Sheet '{sheetName}' is now visible",
                        "Users can see and interact with the sheet"
                    ],
                    SheetVisibility.Hidden => 
                    [
                        $"Sheet '{sheetName}' is now hidden",
                        "Users can unhide via Excel UI (right-click tabs)",
                        "Use 'show' to make visible again"
                    ],
                    SheetVisibility.VeryHidden => 
                    [
                        $"Sheet '{sheetName}' is now very hidden",
                        "Only code can unhide this sheet (protected from users)",
                        "Use 'show' action to make visible when needed"
                    ],
                    _ => []
                };
                
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
                ComUtilities.Release(ref sheet);
            }
        });
    }
    
    /// <inheritdoc />
    public async Task<SheetVisibilityResult> GetVisibilityAsync(IExcelBatch batch, string sheetName)
    {
        var result = new SheetVisibilityResult { FilePath = batch.WorkbookPath };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }
                
                int visibilityValue = Convert.ToInt32(sheet.Visible);
                result.Visibility = (SheetVisibility)visibilityValue;
                result.VisibilityName = result.Visibility.ToString();
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
                ComUtilities.Release(ref sheet);
            }
        });
    }
    
    /// <inheritdoc />
    public async Task<OperationResult> ShowAsync(IExcelBatch batch, string sheetName)
    {
        return await SetVisibilityAsync(batch, sheetName, SheetVisibility.Visible);
    }
    
    /// <inheritdoc />
    public async Task<OperationResult> HideAsync(IExcelBatch batch, string sheetName)
    {
        return await SetVisibilityAsync(batch, sheetName, SheetVisibility.Hidden);
    }
    
    /// <inheritdoc />
    public async Task<OperationResult> VeryHideAsync(IExcelBatch batch, string sheetName)
    {
        return await SetVisibilityAsync(batch, sheetName, SheetVisibility.VeryHidden);
    }
}

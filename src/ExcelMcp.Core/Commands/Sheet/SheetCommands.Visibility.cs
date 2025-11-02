using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet visibility operations (SetVisibility, GetVisibility, Show, Hide, VeryHide)
/// </summary>
public partial class SheetCommands
{
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
    /// <inheritdoc />
    public async Task<OperationResult> ShowAsync(IExcelBatch batch, string sheetName)
    {
        return await SetVisibilityAsync(batch, sheetName, SheetVisibility.Visible);
    }
    
    /// <inheritdoc />
    /// <inheritdoc />
    public async Task<OperationResult> HideAsync(IExcelBatch batch, string sheetName)
    {
        return await SetVisibilityAsync(batch, sheetName, SheetVisibility.Hidden);
    }
    
    /// <inheritdoc />
    /// <inheritdoc />
    public async Task<OperationResult> VeryHideAsync(IExcelBatch batch, string sheetName)
    {
        return await SetVisibilityAsync(batch, sheetName, SheetVisibility.VeryHidden);
    }
}


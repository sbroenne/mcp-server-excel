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
    public OperationResult SetVisibility(IExcelBatch batch, string sheetName, SheetVisibility visibility)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-visibility" };

        return batch.Execute((ctx, ct) =>
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
    public SheetVisibilityResult GetVisibility(IExcelBatch batch, string sheetName)
    {
        var result = new SheetVisibilityResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
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
    public OperationResult Show(IExcelBatch batch, string sheetName)
    {
        return SetVisibility(batch, sheetName, SheetVisibility.Visible);
    }

    /// <inheritdoc />
    /// <inheritdoc />
    public OperationResult Hide(IExcelBatch batch, string sheetName)
    {
        return SetVisibility(batch, sheetName, SheetVisibility.Hidden);
    }

    /// <inheritdoc />
    /// <inheritdoc />
    public OperationResult VeryHide(IExcelBatch batch, string sheetName)
    {
        return SetVisibility(batch, sheetName, SheetVisibility.VeryHidden);
    }
}



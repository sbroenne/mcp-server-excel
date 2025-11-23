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
    public void SetVisibility(IExcelBatch batch, string sheetName, SheetVisibility visibility)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found");
                }

                // Set visibility using the enum value (maps to XlSheetVisibility)
                sheet.Visible = (int)visibility;
                return 0;
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
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found");
                }

                int visibilityValue = Convert.ToInt32(sheet.Visible);
                result.Visibility = (SheetVisibility)visibilityValue;
                result.VisibilityName = result.Visibility.ToString();
                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public void Show(IExcelBatch batch, string sheetName)
    {
        SetVisibility(batch, sheetName, SheetVisibility.Visible);
    }

    /// <inheritdoc />
    public void Hide(IExcelBatch batch, string sheetName)
    {
        SetVisibility(batch, sheetName, SheetVisibility.Hidden);
    }

    /// <inheritdoc />
    public void VeryHide(IExcelBatch batch, string sheetName)
    {
        SetVisibility(batch, sheetName, SheetVisibility.VeryHidden);
    }
}



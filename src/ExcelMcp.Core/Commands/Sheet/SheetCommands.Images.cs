using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet image operations.
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult AddImage(IExcelBatch batch, string sheetName, string imagePath, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            dynamic? pictures = null;
            dynamic? picture = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                cell = sheet.Range[cellAddress];
                if (cell == null)
                {
                    throw new InvalidOperationException($"Cell '{cellAddress}' could not be resolved.");
                }

                pictures = sheet.Pictures(Type.Missing);
                picture = pictures.Insert(imagePath);
                picture.Left = cell.Left;
                picture.Top = cell.Top;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref picture);
                ComUtilities.Release(ref pictures);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public WorksheetImageCountResult GetImageCount(IExcelBatch batch, string sheetName)
    {
        var result = new WorksheetImageCountResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            dynamic? pictures = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                pictures = sheet.Pictures(Type.Missing);
                result.ImageCount = pictures.Count;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref pictures);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

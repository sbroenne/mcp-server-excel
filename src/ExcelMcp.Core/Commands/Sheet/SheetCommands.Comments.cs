using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Legacy worksheet cell note operations through Excel's Comment COM API.
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult SetComment(IExcelBatch batch, string sheetName, string cellAddress, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            dynamic? comment = null;
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

                comment = cell.Comment;
                if (comment == null)
                {
                    comment = cell.AddComment();
                }

                comment.Text(text ?? string.Empty);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref comment);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SheetCommentResult GetComment(IExcelBatch batch, string sheetName, string cellAddress)
    {
        var result = new SheetCommentResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            dynamic? comment = null;
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

                comment = cell.Comment;
                result.HasComment = comment != null;
                if (comment != null)
                {
                    result.Text = comment.Text();
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref comment);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ClearComment(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            dynamic? comment = null;
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

                comment = cell.Comment;
                if (comment != null)
                {
                    comment.Delete();
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref comment);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

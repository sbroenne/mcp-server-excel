using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet shape operations.
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult AddShape(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.TextFrame? textFrame = null;
            Excel.Characters? characters = null;
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

                shapes = sheet.Shapes;

                shape = shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    Convert.ToSingle((object)cell.Left, System.Globalization.CultureInfo.CurrentCulture),
                    Convert.ToSingle((object)cell.Top, System.Globalization.CultureInfo.CurrentCulture),
                    144f,
                    72f);
                textFrame = shape.TextFrame;
                characters = textFrame.Characters();
                characters.Text = "Shape";

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref characters);
                ComUtilities.Release(ref textFrame);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public WorksheetShapeCountResult GetShapeCount(IExcelBatch batch, string sheetName)
    {
        var result = new WorksheetShapeCountResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                shapes = sheet.Shapes;
                result.ShapeCount = shapes.Count;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}

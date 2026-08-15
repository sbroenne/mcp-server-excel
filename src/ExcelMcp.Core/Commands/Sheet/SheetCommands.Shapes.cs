using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Drawing;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

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
            dynamic? shapes = null;
            dynamic? shape = null;
            dynamic? textFrame = null;
            dynamic? characters = null;
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

                // PIA gap: AddShape requires Office.Core.MsoAutoShapeType, while this project intentionally has no office.dll reference.
                dynamic lateBoundShapes = (dynamic)(object)shapes;
                shape = lateBoundShapes.AddShape(
                    (int)DrawingShapeType.Rectangle,
                    Convert.ToSingle(cell.Left),
                    Convert.ToSingle(cell.Top),
                    144f,
                    72f);
                textFrame = shape.TextFrame;
                characters = textFrame.Characters;
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
            dynamic? shapes = null;
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

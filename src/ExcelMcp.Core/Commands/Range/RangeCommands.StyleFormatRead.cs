using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    internal static CellFormatSnapshot ReadStyleFormat(Excel.Style style, CancellationToken ct)
    {
        Excel.Font? font = null;
        Excel.Interior? fill = null;
        Excel.Borders? borders = null;
        try
        {
            font = style.Font;
            fill = style.Interior;
            borders = style.Borders;
            List<string> mixed = [];
            var (fontFormat, fillFormat, borderFormats) = ReadVisualFormat(font, fill, borders, mixed, ct, styleDefinition: true);
            return new CellFormatSnapshot
            {
                Font = fontFormat,
                Fill = fillFormat,
                Borders = borderFormats,
                NumberFormat = style.NumberFormat,
                HorizontalAlignment = ReadFormatValue<int>(style.HorizontalAlignment, "horizontalAlignment", mixed),
                VerticalAlignment = ReadFormatValue<int>(style.VerticalAlignment, "verticalAlignment", mixed),
                WrapText = ReadFormatValue<bool>(style.WrapText, "wrapText", mixed),
                ShrinkToFit = ReadFormatValue<bool>(style.ShrinkToFit, "shrinkToFit", mixed),
                AddIndent = ReadFormatValue<bool>(style.AddIndent, "addIndent", mixed),
                IndentLevel = ReadFormatValue<int>(ComUtilities.ReadStyleIndentLevel(style), "indentLevel", mixed),
                Orientation = ReadFormatValue<int>(style.Orientation, "orientation", mixed),
                ReadingOrder = ReadFormatValue<int>(style.ReadingOrder, "readingOrder", mixed),
                Locked = ReadFormatValue<bool>(style.Locked, "locked", mixed),
                FormulaHidden = ReadFormatValue<bool>(style.FormulaHidden, "formulaHidden", mixed),
                StyleName = style.Name,
                MixedFields = mixed
            };
        }
        finally
        {
            ComUtilities.Release(ref borders);
            ComUtilities.Release(ref fill);
            ComUtilities.Release(ref font);
        }
    }
}

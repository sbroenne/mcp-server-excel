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
            // PIA gap: Style value-type getters can return native null for excluded/unset components.
            dynamic native = style;
            return new CellFormatSnapshot
            {
                Font = fontFormat,
                Fill = fillFormat,
                Borders = borderFormats,
                NumberFormat = style.NumberFormat,
                HorizontalAlignment = ReadFormatValue<int>((object)native.HorizontalAlignment, "horizontalAlignment", mixed),
                VerticalAlignment = ReadFormatValue<int>((object)native.VerticalAlignment, "verticalAlignment", mixed),
                WrapText = ReadFormatValue<bool>((object)native.WrapText, "wrapText", mixed),
                ShrinkToFit = ReadFormatValue<bool>((object)native.ShrinkToFit, "shrinkToFit", mixed),
                AddIndent = ReadFormatValue<bool>((object)native.AddIndent, "addIndent", mixed),
                IndentLevel = ReadFormatValue<int>((object)native.IndentLevel, "indentLevel", mixed),
                Orientation = ReadFormatValue<int>((object)native.Orientation, "orientation", mixed),
                ReadingOrder = ReadFormatValue<int>((object)native.ReadingOrder, "readingOrder", mixed),
                Locked = ReadFormatValue<bool>((object)native.Locked, "locked", mixed),
                FormulaHidden = ReadFormatValue<bool>((object)native.FormulaHidden, "formulaHidden", mixed),
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

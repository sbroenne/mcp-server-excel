using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet tab color operations (SetTabColor, GetTabColor, ClearTabColor)
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult SetTabColor(IExcelBatch batch, string sheetName, int red, int green, int blue)
    {
        // Validate RGB values
        if (red < 0 || red > 255 || green < 0 || green > 255 || blue < 0 || blue > 255)
        {
            throw new ArgumentException("RGB values must be between 0 and 255");
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? tab = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                // Convert RGB to BGR format (Excel's color format)
                // BGR = (Blue << 16) | (Green << 8) | Red
                int bgrColor = (blue << 16) | (green << 8) | red;

                tab = sheet.Tab;
                tab.Color = bgrColor;
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref tab);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    /// <inheritdoc />
    public TabColorResult GetTabColor(IExcelBatch batch, string sheetName)
    {
        var result = new TabColorResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? tab = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
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
                if (colorValue is null or (dynamic?)0)
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
            finally
            {
                ComUtilities.Release(ref tab);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ClearTabColor(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? tab = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                tab = sheet.Tab;
                // Set ColorIndex to xlColorIndexNone (-4142) to clear color
                tab.ColorIndex = -4142; // xlColorIndexNone
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref tab);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}




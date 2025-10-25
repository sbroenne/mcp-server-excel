using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands implementation
/// </summary>
public class CellCommands : ICellCommands
{
    /// <inheritdoc />
    public CellValueResult GetValue(string filePath, string sheetName, string cellAddress)
    {
        if (!File.Exists(filePath))
        {
            return new CellValueResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                CellAddress = cellAddress
            };
        }

        var result = new CellValueResult
        {
            FilePath = filePath,
            CellAddress = $"{sheetName}!{cellAddress}"
        };

        ExcelSession.Execute(filePath, false, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                cell = sheet.Range[cellAddress];
                result.Value = cell.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
                result.Formula = cell.Formula;
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetValue(string filePath, string sheetName, string cellAddress, string value)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-value"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-value"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                cell = sheet.Range[cellAddress];

                // Try to parse as number, otherwise set as text
                if (double.TryParse(value, out double numValue))
                {
                    cell.Value2 = numValue;
                }
                else if (bool.TryParse(value, out bool boolValue))
                {
                    cell.Value2 = boolValue;
                }
                else
                {
                    cell.Value2 = value;
                }

                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public CellValueResult GetFormula(string filePath, string sheetName, string cellAddress)
    {
        if (!File.Exists(filePath))
        {
            return new CellValueResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                CellAddress = cellAddress
            };
        }

        var result = new CellValueResult
        {
            FilePath = filePath,
            CellAddress = $"{sheetName}!{cellAddress}"
        };

        ExcelSession.Execute(filePath, false, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                cell = sheet.Range[cellAddress];
                result.Formula = cell.Formula ?? "";
                result.Value = cell.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetFormula(string filePath, string sheetName, string cellAddress, string formula)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-formula"
            };
        }

        // Ensure formula starts with =
        if (!formula.StartsWith("="))
        {
            formula = "=" + formula;
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-formula"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                cell = sheet.Range[cellAddress];
                cell.Formula = formula;

                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetBackgroundColor(string filePath, string sheetName, string cellAddress, string color)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-background-color"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-background-color"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? interior = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                range = sheet.Range[cellAddress];
                interior = range.Interior;
                
                // Parse color - support hex (#RRGGBB), RGB (r,g,b), or color name
                if (color.StartsWith("#"))
                {
                    // Hex color
                    var rgb = Convert.ToInt32(color.Substring(1), 16);
                    interior.Color = rgb;
                }
                else if (color.Contains(","))
                {
                    // RGB format
                    var parts = color.Split(',');
                    if (parts.Length == 3 && 
                        int.TryParse(parts[0].Trim(), out int r) &&
                        int.TryParse(parts[1].Trim(), out int g) &&
                        int.TryParse(parts[2].Trim(), out int b))
                    {
                        interior.Color = r + (g * 256) + (b * 256 * 256);
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = "Invalid RGB format. Use 'r,g,b' where each value is 0-255";
                        return 1;
                    }
                }
                else
                {
                    // Try to parse as VBA color constant or number
                    if (int.TryParse(color, out int colorInt))
                    {
                        interior.Color = colorInt;
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Invalid color format: {color}. Use hex (#RRGGBB), RGB (r,g,b), or color number";
                        return 1;
                    }
                }

                result.Success = true;
                result.WorkflowHint = $"Background color set for {cellAddress}";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetFontColor(string filePath, string sheetName, string cellAddress, string color)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-font-color"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-font-color"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                range = sheet.Range[cellAddress];
                font = range.Font;
                
                // Parse color - same logic as background color
                if (color.StartsWith("#"))
                {
                    var rgb = Convert.ToInt32(color.Substring(1), 16);
                    font.Color = rgb;
                }
                else if (color.Contains(","))
                {
                    var parts = color.Split(',');
                    if (parts.Length == 3 && 
                        int.TryParse(parts[0].Trim(), out int r) &&
                        int.TryParse(parts[1].Trim(), out int g) &&
                        int.TryParse(parts[2].Trim(), out int b))
                    {
                        font.Color = r + (g * 256) + (b * 256 * 256);
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = "Invalid RGB format. Use 'r,g,b' where each value is 0-255";
                        return 1;
                    }
                }
                else
                {
                    if (int.TryParse(color, out int colorInt))
                    {
                        font.Color = colorInt;
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Invalid color format: {color}. Use hex (#RRGGBB), RGB (r,g,b), or color number";
                        return 1;
                    }
                }

                result.Success = true;
                result.WorkflowHint = $"Font color set for {cellAddress}";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetFont(string filePath, string sheetName, string cellAddress, string? fontName = null, int? fontSize = null, bool? bold = null, bool? italic = null, bool? underline = null)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-font"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-font"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                range = sheet.Range[cellAddress];
                font = range.Font;
                
                if (fontName != null) font.Name = fontName;
                if (fontSize.HasValue) font.Size = fontSize.Value;
                if (bold.HasValue) font.Bold = bold.Value;
                if (italic.HasValue) font.Italic = italic.Value;
                if (underline.HasValue) font.Underline = underline.Value ? 2 : -4142; // xlUnderlineStyleSingle = 2, xlUnderlineStyleNone = -4142

                result.Success = true;
                result.WorkflowHint = $"Font properties set for {cellAddress}";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetBorder(string filePath, string sheetName, string cellAddress, string borderStyle, string? borderColor = null)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-border"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-border"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? borders = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                range = sheet.Range[cellAddress];
                borders = range.Borders;
                
                // Map border style to Excel constant
                // xlContinuous = 1, xlDash = -4115, xlDot = -4118, xlDouble = -4119, xlNone = -4142
                int lineStyle = borderStyle.ToLowerInvariant() switch
                {
                    "thin" or "continuous" => 1,
                    "dash" or "dashed" => -4115,
                    "dot" or "dotted" => -4118,
                    "double" => -4119,
                    "none" => -4142,
                    _ => 1
                };

                // Apply to all borders
                for (int i = 7; i <= 12; i++) // xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10, xlInsideVertical=11, xlInsideHorizontal=12
                {
                    dynamic? border = null;
                    try
                    {
                        border = borders.Item(i);
                        border.LineStyle = lineStyle;
                        if (borderColor != null)
                        {
                            if (borderColor.StartsWith("#"))
                            {
                                var rgb = Convert.ToInt32(borderColor.Substring(1), 16);
                                border.Color = rgb;
                            }
                            else if (int.TryParse(borderColor, out int colorInt))
                            {
                                border.Color = colorInt;
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref border);
                    }
                }

                result.Success = true;
                result.WorkflowHint = $"Border set for {cellAddress}";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetNumberFormat(string filePath, string sheetName, string cellAddress, string format)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-number-format"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-number-format"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                range = sheet.Range[cellAddress];
                range.NumberFormat = format;

                result.Success = true;
                result.WorkflowHint = $"Number format set to '{format}' for {cellAddress}";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult SetAlignment(string filePath, string sheetName, string cellAddress, string? horizontal = null, string? vertical = null, bool? wrapText = null)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-alignment"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-alignment"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                range = sheet.Range[cellAddress];
                
                if (horizontal != null)
                {
                    // xlLeft = -4131, xlCenter = -4108, xlRight = -4152, xlJustify = -4130
                    int hAlign = horizontal.ToLowerInvariant() switch
                    {
                        "left" => -4131,
                        "center" => -4108,
                        "right" => -4152,
                        "justify" => -4130,
                        _ => -4131
                    };
                    range.HorizontalAlignment = hAlign;
                }

                if (vertical != null)
                {
                    // xlTop = -4160, xlCenter = -4108, xlBottom = -4107
                    int vAlign = vertical.ToLowerInvariant() switch
                    {
                        "top" => -4160,
                        "center" or "middle" => -4108,
                        "bottom" => -4107,
                        _ => -4107
                    };
                    range.VerticalAlignment = vAlign;
                }

                if (wrapText.HasValue)
                {
                    range.WrapText = wrapText.Value;
                }

                result.Success = true;
                result.WorkflowHint = $"Alignment set for {cellAddress}";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult ClearFormatting(string filePath, string sheetName, string cellAddress)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "clear-formatting"
            };
        }

        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "clear-formatting"
        };

        ExcelSession.Execute(filePath, true, (excel, workbook) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return 1;
                }

                range = sheet.Range[cellAddress];
                range.ClearFormats();

                result.Success = true;
                result.WorkflowHint = $"Formatting cleared for {cellAddress}";
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        return result;
    }
}

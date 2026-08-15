using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

public sealed partial class DrawingCommands
{
    /// <inheritdoc />
    public DrawingObjectListResult ListObjects(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;
                var result = new DrawingObjectListResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };

                for (var index = 1; index <= shapes.Count; index++)
                {
                    Excel.Shape? shape = null;
                    try
                    {
                        shape = shapes.Item(index);
                        result.DrawingObjects.Add(ReadDrawingObject(shape, sheetName));
                    }
                    finally
                    {
                        ComUtilities.Release(ref shape);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectResult GetObject(IExcelBatch batch, string sheetName, string objectName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;
                shape = FindShape(shapes, objectName)
                    ?? throw new InvalidOperationException($"Drawing object '{objectName}' not found on sheet '{sheetName}'.");
                return CreateDrawingObjectResult(batch.WorkbookPath, ReadDrawingObject(shape, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectResult AddImage(
        IExcelBatch batch,
        string sheetName,
        string imagePath,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 120,
        double height = 80,
        bool lockAspectRatio = true)
    {
        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"Image file not found: {imagePath}", imagePath);
        }

        ValidateGeometry(width, height);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;

                // PIA gap: AddPicture requires Office.Core.MsoTriState, while this project intentionally has no office.dll reference.
                dynamic lateBoundShapes = (dynamic)(object)shapes;
                shape = (Excel.Shape)lateBoundShapes.AddPicture(
                    Path.GetFullPath(imagePath),
                    0,
                    -1,
                    Convert.ToSingle(left),
                    Convert.ToSingle(top),
                    Convert.ToSingle(width),
                    Convert.ToSingle(height));

                ApplyName(shape, name);
                SetLockAspectRatio(shape, lockAspectRatio);
                return CreateDrawingObjectResult(batch.WorkbookPath, ReadDrawingObject(shape, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectResult AddShape(
        IExcelBatch batch,
        string sheetName,
        DrawingShapeType shapeType = DrawingShapeType.Rectangle,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 120,
        double height = 60,
        string? text = null,
        string? fillColor = null,
        string? lineColor = null,
        double? lineWeight = null)
    {
        ValidateGeometry(width, height);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;

                // PIA gap: AddShape requires Office.Core.MsoAutoShapeType, while this project intentionally has no office.dll reference.
                dynamic lateBoundShapes = (dynamic)(object)shapes;
                shape = (Excel.Shape)lateBoundShapes.AddShape(
                    (int)shapeType,
                    Convert.ToSingle(left),
                    Convert.ToSingle(top),
                    Convert.ToSingle(width),
                    Convert.ToSingle(height));

                ApplyName(shape, name);
                ApplyObjectFormatting(shape, text, null, null, fillColor, lineColor, lineWeight);
                return CreateDrawingObjectResult(batch.WorkbookPath, ReadDrawingObject(shape, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectResult AddTextBox(
        IExcelBatch batch,
        string sheetName,
        string text,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 180,
        double height = 50,
        double? fontSize = null,
        string? fontColor = null,
        string? fillColor = null,
        string? lineColor = null)
    {
        ValidateGeometry(width, height);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;

                // PIA gap: AddTextbox requires Office.Core.MsoTextOrientation, while this project intentionally has no office.dll reference.
                dynamic lateBoundShapes = (dynamic)(object)shapes;
                shape = (Excel.Shape)lateBoundShapes.AddTextbox(
                    1,
                    Convert.ToSingle(left),
                    Convert.ToSingle(top),
                    Convert.ToSingle(width),
                    Convert.ToSingle(height));

                ApplyName(shape, name);
                ApplyObjectFormatting(shape, text, fontSize, fontColor, fillColor, lineColor, null);
                return CreateDrawingObjectResult(batch.WorkbookPath, ReadDrawingObject(shape, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectResult AddConnector(
        IExcelBatch batch,
        string sheetName,
        DrawingConnectorType connectorType = DrawingConnectorType.Straight,
        double beginX = 20,
        double beginY = 20,
        double endX = 140,
        double endY = 20,
        string? name = null,
        string? lineColor = null,
        double? lineWeight = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;

                // PIA gap: AddConnector requires Office.Core.MsoConnectorType, while this project intentionally has no office.dll reference.
                dynamic lateBoundShapes = (dynamic)(object)shapes;
                shape = (Excel.Shape)lateBoundShapes.AddConnector(
                    (int)connectorType,
                    Convert.ToSingle(beginX),
                    Convert.ToSingle(beginY),
                    Convert.ToSingle(endX),
                    Convert.ToSingle(endY));

                ApplyName(shape, name);
                ApplyLineFormatting(shape, lineColor, lineWeight);
                return CreateDrawingObjectResult(batch.WorkbookPath, ReadDrawingObject(shape, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectResult AddFormControl(
        IExcelBatch batch,
        string sheetName,
        DrawingFormControlType controlType = DrawingFormControlType.Button,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 120,
        double height = 24,
        string? text = null,
        string? linkedCell = null,
        string? inputRange = null)
    {
        ValidateGeometry(width, height);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.ControlFormat? controlFormat = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;
                shape = shapes.AddFormControl(
                    (Excel.XlFormControl)controlType,
                    Convert.ToInt32(left),
                    Convert.ToInt32(top),
                    Convert.ToInt32(width),
                    Convert.ToInt32(height));

                ApplyName(shape, name);
                if (text != null)
                {
                    SetText(shape, text, null, null);
                }

                if (linkedCell != null || inputRange != null)
                {
                    controlFormat = shape.ControlFormat;
                    if (linkedCell != null)
                    {
                        controlFormat.LinkedCell = linkedCell;
                    }

                    if (inputRange != null)
                    {
                        controlFormat.ListFillRange = inputRange;
                    }
                }

                return CreateDrawingObjectResult(batch.WorkbookPath, ReadDrawingObject(shape, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref controlFormat);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectResult UpdateObject(
        IExcelBatch batch,
        string sheetName,
        string objectName,
        string? newName = null,
        double? left = null,
        double? top = null,
        double? width = null,
        double? height = null,
        double? rotation = null,
        string? text = null,
        double? fontSize = null,
        string? fontColor = null,
        string? fillColor = null,
        string? lineColor = null,
        double? lineWeight = null,
        bool? visible = null,
        bool? locked = null,
        int? placement = null,
        string? alternativeText = null,
        string? linkedCell = null,
        string? inputRange = null)
    {
        if (width.HasValue && width <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "Width must be greater than zero.");
        }

        if (height.HasValue && height <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(height), "Height must be greater than zero.");
        }

        if (placement.HasValue && placement is < 1 or > 3)
        {
            throw new ArgumentOutOfRangeException(nameof(placement), "Placement must be 1, 2, or 3.");
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.ControlFormat? controlFormat = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;
                shape = FindShape(shapes, objectName)
                    ?? throw new InvalidOperationException($"Drawing object '{objectName}' not found on sheet '{sheetName}'.");

                if (newName != null) shape.Name = newName;
                if (left.HasValue) shape.Left = Convert.ToSingle(left.Value);
                if (top.HasValue) shape.Top = Convert.ToSingle(top.Value);
                if (width.HasValue) shape.Width = Convert.ToSingle(width.Value);
                if (height.HasValue) shape.Height = Convert.ToSingle(height.Value);
                if (rotation.HasValue) shape.Rotation = Convert.ToSingle(rotation.Value);
                if (locked.HasValue) shape.Locked = locked.Value;
                if (placement.HasValue) shape.Placement = (Excel.XlPlacement)placement.Value;
                if (alternativeText != null) shape.AlternativeText = alternativeText;
                if (visible.HasValue) SetVisible(shape, visible.Value);

                ApplyObjectFormatting(shape, text, fontSize, fontColor, fillColor, lineColor, lineWeight);

                if (linkedCell != null || inputRange != null)
                {
                    if (ReadKind(shape) != DrawingObjectKind.FormControl)
                    {
                        throw new InvalidOperationException("linkedCell and inputRange apply only to worksheet Forms controls.");
                    }

                    controlFormat = shape.ControlFormat;
                    if (linkedCell != null) controlFormat.LinkedCell = linkedCell;
                    if (inputRange != null) controlFormat.ListFillRange = inputRange;
                }

                return CreateDrawingObjectResult(batch.WorkbookPath, ReadDrawingObject(shape, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref controlFormat);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteObject(IExcelBatch batch, string sheetName, string objectName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                shapes = sheet.Shapes;
                shape = FindShape(shapes, objectName)
                    ?? throw new InvalidOperationException($"Drawing object '{objectName}' not found on sheet '{sheetName}'.");
                shape.Delete();
                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "delete-object"
                };
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static DrawingObjectResult CreateDrawingObjectResult(string filePath, DrawingObjectInfo drawingObject)
    {
        return new DrawingObjectResult
        {
            Success = true,
            FilePath = filePath,
            DrawingObject = drawingObject
        };
    }

    private static Excel.Worksheet GetSheet(Excel.Workbook workbook, string sheetName)
    {
        return ComUtilities.FindSheet(workbook, sheetName)
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
    }

    private static Excel.Shape? FindShape(Excel.Shapes shapes, string objectName)
    {
        for (var index = 1; index <= shapes.Count; index++)
        {
            Excel.Shape? shape = null;
            try
            {
                shape = shapes.Item(index);
                if (string.Equals(shape.Name, objectName, StringComparison.OrdinalIgnoreCase))
                {
                    var found = shape;
                    shape = null;
                    return found;
                }
            }
            finally
            {
                ComUtilities.Release(ref shape);
            }
        }

        return null;
    }

    private static DrawingObjectInfo ReadDrawingObject(Excel.Shape shape, string sheetName)
    {
        var kind = ReadKind(shape);
        var result = new DrawingObjectInfo
        {
            Name = shape.Name,
            SheetName = sheetName,
            Kind = kind,
            Left = shape.Left,
            Top = shape.Top,
            Width = shape.Width,
            Height = shape.Height,
            Rotation = shape.Rotation,
            Visible = ReadVisible(shape),
            Locked = shape.Locked,
            Placement = Convert.ToInt32(shape.Placement, System.Globalization.CultureInfo.InvariantCulture),
            AlternativeText = NullIfEmpty(shape.AlternativeText),
            FillColor = ReadFillColor(shape),
            LineColor = ReadLineColor(shape),
            LineWeight = ReadLineWeight(shape)
        };

        if (kind == DrawingObjectKind.AutoShape)
        {
            result.ShapeType = ReadShapeType(shape);
        }
        else if (kind == DrawingObjectKind.Connector)
        {
            result.ConnectorType = ReadConnectorType(shape);
        }
        else if (kind == DrawingObjectKind.FormControl)
        {
            var controlType = (DrawingFormControlType)shape.FormControlType;
            result.FormControlType = controlType;
            ReadControlProperties(shape, controlType, result);
        }

        ReadTextProperties(shape, result);
        return result;
    }

    private static DrawingObjectKind ReadKind(Excel.Shape shape)
    {
        // PIA gap: Shape.Type and Shape.Connector return Office.Core enum types, while this project intentionally has no office.dll reference.
        dynamic lateBoundShape = (dynamic)(object)shape;
        if (Convert.ToInt32(lateBoundShape.Connector) != 0)
        {
            return DrawingObjectKind.Connector;
        }

        var shapeType = Convert.ToInt32(lateBoundShape.Type);
        return shapeType switch
        {
            1 => DrawingObjectKind.AutoShape,
            8 => DrawingObjectKind.FormControl,
            11 or 13 => DrawingObjectKind.Image,
            17 => DrawingObjectKind.TextBox,
            _ => DrawingObjectKind.Other
        };
    }

    private static DrawingShapeType? ReadShapeType(Excel.Shape shape)
    {
        // PIA gap: Shape.AutoShapeType returns Office.Core.MsoAutoShapeType.
        dynamic lateBoundShape = (dynamic)(object)shape;
        var value = Convert.ToInt32(lateBoundShape.AutoShapeType);
        return Enum.IsDefined(typeof(DrawingShapeType), value) ? (DrawingShapeType)value : null;
    }

    private static DrawingConnectorType? ReadConnectorType(Excel.Shape shape)
    {
        dynamic? connectorFormat = null;
        try
        {
            // PIA gap: Shape.ConnectorFormat returns an Office.Core object.
            dynamic lateBoundShape = (dynamic)(object)shape;
            connectorFormat = lateBoundShape.ConnectorFormat;
            var value = Convert.ToInt32(connectorFormat.Type);
            return Enum.IsDefined(typeof(DrawingConnectorType), value) ? (DrawingConnectorType)value : null;
        }
        finally
        {
            ComUtilities.Release(ref connectorFormat);
        }
    }

    private static void ApplyName(Excel.Shape shape, string? name)
    {
        if (!string.IsNullOrWhiteSpace(name))
        {
            shape.Name = name;
        }
    }

    private static void ApplyObjectFormatting(
        Excel.Shape shape,
        string? text,
        double? fontSize,
        string? fontColor,
        string? fillColor,
        string? lineColor,
        double? lineWeight)
    {
        if (text != null || fontSize.HasValue || fontColor != null)
        {
            SetText(shape, text, fontSize, fontColor);
        }

        if (fillColor != null)
        {
            SetFillColor(shape, fillColor);
        }

        ApplyLineFormatting(shape, lineColor, lineWeight);
    }

    private static void SetText(Excel.Shape shape, string? text, double? fontSize, string? fontColor)
    {
        Excel.TextFrame? textFrame = null;
        Excel.Characters? characters = null;
        Excel.Font? font = null;
        try
        {
            textFrame = shape.TextFrame;
            characters = textFrame.Characters();
            if (text != null) characters.Text = text;
            if (fontSize.HasValue || fontColor != null)
            {
                font = characters.Font;
                if (fontSize.HasValue) font.Size = fontSize.Value;
                if (fontColor != null) font.Color = ParseColor(fontColor);
            }
        }
        finally
        {
            ComUtilities.Release(ref font);
            ComUtilities.Release(ref characters);
            ComUtilities.Release(ref textFrame);
        }
    }

    private static void ReadTextProperties(Excel.Shape shape, DrawingObjectInfo result)
    {
        Excel.TextFrame? textFrame = null;
        Excel.Characters? characters = null;
        Excel.Font? font = null;
        try
        {
            textFrame = shape.TextFrame;
            characters = textFrame.Characters();
            result.Text = NullIfEmpty(characters.Text);
            if (result.Text == null)
            {
                return;
            }

            font = characters.Font;
            object? fontSize = font.Size;
            object? fontColor = font.Color;
            result.FontSize = fontSize is DBNull
                ? null
                : Convert.ToDouble(fontSize, System.Globalization.CultureInfo.InvariantCulture);
            result.FontColor = fontColor is DBNull
                ? null
                : FormatColor(Convert.ToInt32(fontColor, System.Globalization.CultureInfo.InvariantCulture));
        }
        catch (COMException)
        {
            result.Text = null;
            result.FontSize = null;
            result.FontColor = null;
        }
        finally
        {
            ComUtilities.Release(ref font);
            ComUtilities.Release(ref characters);
            ComUtilities.Release(ref textFrame);
        }
    }

    private static void ReadControlProperties(
        Excel.Shape shape,
        DrawingFormControlType controlType,
        DrawingObjectInfo result)
    {
        result.LinkedCell = null;
        result.InputRange = null;

        Excel.ControlFormat? controlFormat = null;
        try
        {
            controlFormat = shape.ControlFormat;
            if (SupportsLinkedCell(controlType))
            {
                result.LinkedCell = NullIfEmpty(controlFormat.LinkedCell);
            }

            if (SupportsInputRange(controlType))
            {
                result.InputRange = NullIfEmpty(controlFormat.ListFillRange);
            }
        }
        finally
        {
            ComUtilities.Release(ref controlFormat);
        }
    }

    private static bool SupportsLinkedCell(DrawingFormControlType controlType)
    {
        return controlType is
            DrawingFormControlType.CheckBox or
            DrawingFormControlType.DropDown or
            DrawingFormControlType.ListBox or
            DrawingFormControlType.OptionButton or
            DrawingFormControlType.ScrollBar or
            DrawingFormControlType.Spinner;
    }

    private static bool SupportsInputRange(DrawingFormControlType controlType)
    {
        return controlType is DrawingFormControlType.DropDown or DrawingFormControlType.ListBox;
    }

    private static void SetFillColor(Excel.Shape shape, string color)
    {
        dynamic? fill = null;
        dynamic? colorFormat = null;
        try
        {
            // PIA gap: Shape.Fill and ColorFormat are Office.Core types.
            dynamic lateBoundShape = (dynamic)(object)shape;
            fill = lateBoundShape.Fill;
            colorFormat = fill.ForeColor;
            fill.Visible = -1;
            colorFormat.RGB = ParseColor(color);
        }
        finally
        {
            ComUtilities.Release(ref colorFormat);
            ComUtilities.Release(ref fill);
        }
    }

    private static string? ReadFillColor(Excel.Shape shape)
    {
        dynamic? fill = null;
        dynamic? colorFormat = null;
        try
        {
            // PIA gap: Shape.Fill and ColorFormat are Office.Core types.
            dynamic lateBoundShape = (dynamic)(object)shape;
            fill = lateBoundShape.Fill;
            if (Convert.ToInt32(fill.Visible) == 0)
            {
                return null;
            }

            colorFormat = fill.ForeColor;
            return FormatColor(Convert.ToInt32(colorFormat.RGB));
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            ComUtilities.Release(ref colorFormat);
            ComUtilities.Release(ref fill);
        }
    }

    private static void ApplyLineFormatting(Excel.Shape shape, string? lineColor, double? lineWeight)
    {
        if (lineColor == null && !lineWeight.HasValue)
        {
            return;
        }

        dynamic? line = null;
        dynamic? colorFormat = null;
        try
        {
            // PIA gap: Shape.Line and ColorFormat are Office.Core types.
            dynamic lateBoundShape = (dynamic)(object)shape;
            line = lateBoundShape.Line;
            line.Visible = -1;
            if (lineColor != null)
            {
                colorFormat = line.ForeColor;
                colorFormat.RGB = ParseColor(lineColor);
            }
            if (lineWeight.HasValue)
            {
                line.Weight = Convert.ToSingle(lineWeight.Value);
            }
        }
        finally
        {
            ComUtilities.Release(ref colorFormat);
            ComUtilities.Release(ref line);
        }
    }

    private static string? ReadLineColor(Excel.Shape shape)
    {
        dynamic? line = null;
        dynamic? colorFormat = null;
        try
        {
            // PIA gap: Shape.Line and ColorFormat are Office.Core types.
            dynamic lateBoundShape = (dynamic)(object)shape;
            line = lateBoundShape.Line;
            if (Convert.ToInt32(line.Visible) == 0)
            {
                return null;
            }

            colorFormat = line.ForeColor;
            return FormatColor(Convert.ToInt32(colorFormat.RGB));
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            ComUtilities.Release(ref colorFormat);
            ComUtilities.Release(ref line);
        }
    }

    private static double? ReadLineWeight(Excel.Shape shape)
    {
        dynamic? line = null;
        try
        {
            // PIA gap: Shape.Line returns an Office.Core.LineFormat object.
            dynamic lateBoundShape = (dynamic)(object)shape;
            line = lateBoundShape.Line;
            return Convert.ToDouble(line.Weight);
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            ComUtilities.Release(ref line);
        }
    }

    private static void SetVisible(Excel.Shape shape, bool visible)
    {
        // PIA gap: Shape.Visible uses Office.Core.MsoTriState.
        dynamic lateBoundShape = (dynamic)(object)shape;
        lateBoundShape.Visible = visible ? -1 : 0;
    }

    private static bool ReadVisible(Excel.Shape shape)
    {
        // PIA gap: Shape.Visible uses Office.Core.MsoTriState.
        dynamic lateBoundShape = (dynamic)(object)shape;
        return Convert.ToInt32(lateBoundShape.Visible) != 0;
    }

    private static void SetLockAspectRatio(Excel.Shape shape, bool lockAspectRatio)
    {
        // PIA gap: Shape.LockAspectRatio uses Office.Core.MsoTriState.
        dynamic lateBoundShape = (dynamic)(object)shape;
        lateBoundShape.LockAspectRatio = lockAspectRatio ? -1 : 0;
    }

    private static int ParseColor(string color)
    {
        var value = color.Trim();
        if (value.StartsWith('#'))
        {
            value = value[1..];
        }

        if (value.Length != 6 || !int.TryParse(value, System.Globalization.NumberStyles.HexNumber, null, out var rgb))
        {
            throw new ArgumentException($"Invalid color '{color}'. Use #RRGGBB.", nameof(color));
        }

        var red = (rgb >> 16) & 0xFF;
        var green = (rgb >> 8) & 0xFF;
        var blue = rgb & 0xFF;
        return red | (green << 8) | (blue << 16);
    }

    private static string FormatColor(int oleColor)
    {
        var red = oleColor & 0xFF;
        var green = (oleColor >> 8) & 0xFF;
        var blue = (oleColor >> 16) & 0xFF;
        return $"#{red:X2}{green:X2}{blue:X2}";
    }

    private static string? NullIfEmpty(string? value) => string.IsNullOrEmpty(value) ? null : value;

    private static void ValidateGeometry(double width, double height)
    {
        if (width <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "Width must be greater than zero.");
        }

        if (height <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(height), "Height must be greater than zero.");
        }
    }
}

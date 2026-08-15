using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>
/// Worksheet drawing objects and sparklines.
///
/// OBJECTS: list/read/update/delete images, AutoShapes, text boxes, connectors, and worksheet Forms controls.
/// SHAPE TYPES: common geometric, arrow, and flowchart AutoShapes.
/// FORMATTING: geometry, text, fill/line/font colors, rotation, visibility, locking, placement, and alternative text.
/// FORMS CONTROLS: safe worksheet Forms controls only. linkedCell applies to CheckBox, DropDown, ListBox, OptionButton, ScrollBar, and Spinner; inputRange applies only to DropDown and ListBox. ActiveX/OLE controls and macro assignment are intentionally excluded.
/// SPARKLINES: list/read/create/update/delete line, column, and win/loss sparklines.
/// COLORS: use #RRGGBB hexadecimal values.
/// </summary>
[ServiceCategory("drawing", "Drawing")]
[McpTool("drawing", Title = "Drawing Object Operations", Destructive = true, Category = "structure",
    Description = "Worksheet drawing objects and sparklines. Manage images, AutoShapes, text boxes, connectors, and safe worksheet Forms controls with list/read/update/delete lifecycle and formatting. Add common geometric, arrow, and flowchart AutoShapes. Colors use #RRGGBB. Forms controls exclude ActiveX/OLE and macro assignment. Manage line, column, and win/loss sparklines. ")]
public interface IDrawingCommands
{
    /// <summary>Lists drawing objects on a worksheet.</summary>
    [ServiceAction("list-objects")]
    DrawingObjectListResult ListObjects(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>Reads one drawing object by name.</summary>
    [ServiceAction("get-object")]
    DrawingObjectResult GetObject(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string objectName);

    /// <summary>Adds an embedded image from a local file.</summary>
    [ServiceAction("add-image")]
    DrawingObjectResult AddImage(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string imagePath,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 120,
        double height = 80,
        bool lockAspectRatio = true);

    /// <summary>Adds and formats an Excel AutoShape.</summary>
    [ServiceAction("add-shape")]
    DrawingObjectResult AddShape(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        DrawingShapeType shapeType = DrawingShapeType.Rectangle,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 120,
        double height = 60,
        string? text = null,
        string? fillColor = null,
        string? lineColor = null,
        double? lineWeight = null);

    /// <summary>Adds and formats a text box.</summary>
    [ServiceAction("add-text-box")]
    DrawingObjectResult AddTextBox(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string text,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 180,
        double height = 50,
        double? fontSize = null,
        string? fontColor = null,
        string? fillColor = null,
        string? lineColor = null);

    /// <summary>Adds and formats a connector.</summary>
    [ServiceAction("add-connector")]
    DrawingObjectResult AddConnector(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        DrawingConnectorType connectorType = DrawingConnectorType.Straight,
        double beginX = 20,
        double beginY = 20,
        double endX = 140,
        double endY = 20,
        string? name = null,
        string? lineColor = null,
        double? lineWeight = null);

    /// <summary>Adds a safe worksheet Forms control. linkedCell applies to value controls; inputRange applies only to DropDown and ListBox. ActiveX/OLE controls are not supported.</summary>
    [ServiceAction("add-form-control")]
    DrawingObjectResult AddFormControl(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        DrawingFormControlType controlType = DrawingFormControlType.Button,
        string? name = null,
        double left = 20,
        double top = 20,
        double width = 120,
        double height = 24,
        string? text = null,
        string? linkedCell = null,
        string? inputRange = null);

    /// <summary>Updates geometry, formatting, text, accessibility, or Forms-control bindings.</summary>
    [ServiceAction("update-object")]
    DrawingObjectResult UpdateObject(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string objectName,
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
        string? inputRange = null);

    /// <summary>Deletes a drawing object by name.</summary>
    [ServiceAction("delete-object")]
    OperationResult DeleteObject(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string objectName);

    /// <summary>Lists sparkline groups on a worksheet.</summary>
    [ServiceAction("list-sparklines")]
    SparklineListResult ListSparklines(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>Reads the sparkline group at a cell or range.</summary>
    [ServiceAction("get-sparkline")]
    SparklineResult GetSparkline(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string locationRange);

    /// <summary>Adds a line, column, or win/loss sparkline group.</summary>
    [ServiceAction("add-sparkline")]
    SparklineResult AddSparkline(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string sourceRange,
        [RequiredParameter] string locationRange,
        DrawingSparklineType sparklineType = DrawingSparklineType.Line,
        string? lineColor = null,
        bool showMarkers = false);

    /// <summary>Updates a sparkline group's source, type, color, or markers.</summary>
    [ServiceAction("update-sparkline")]
    SparklineResult UpdateSparkline(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string locationRange,
        string? sourceRange = null,
        DrawingSparklineType? sparklineType = null,
        string? lineColor = null,
        bool? showMarkers = null);

    /// <summary>Deletes the sparkline group at a cell or range.</summary>
    [ServiceAction("delete-sparkline")]
    OperationResult DeleteSparkline(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string locationRange);
}

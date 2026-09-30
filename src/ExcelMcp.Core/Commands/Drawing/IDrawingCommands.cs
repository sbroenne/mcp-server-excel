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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the drawing objects or sparklines</param>
    [ServiceAction("list-objects")]
    DrawingObjectListResult ListObjects(IExcelBatch batch, [RequiredParameter] string sheetName);

    /// <summary>Reads one drawing object by name.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the object</param>
    /// <param name="objectName">Existing drawing object's name, as returned by list-objects</param>
    [ServiceAction("get-object")]
    DrawingObjectResult GetObject(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string objectName);

    /// <summary>Adds an embedded image from a local file.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="imagePath">Full path to a readable local image file</param>
    /// <param name="name">Optional name for the new drawing object</param>
    /// <param name="left">Left position in points from the worksheet edge</param>
    /// <param name="top">Top position in points from the worksheet edge</param>
    /// <param name="width">Object width in points</param>
    /// <param name="height">Object height in points</param>
    /// <param name="lockAspectRatio">Keep the image's aspect ratio when resizing</param>
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="shapeType">AutoShape type, such as Rectangle, Oval, or a supported arrow/flowchart shape</param>
    /// <param name="name">Optional object name</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points</param>
    /// <param name="height">Height in points</param>
    /// <param name="text">Text displayed by the shape, text box, or Forms control</param>
    /// <param name="fillColor">Fill color as #RRGGBB</param>
    /// <param name="lineColor">Outline, connector, or sparkline color as #RRGGBB</param>
    /// <param name="lineWeight">Line thickness in points</param>
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="text">Text to display</param>
    /// <param name="name">Optional object name</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points</param>
    /// <param name="height">Height in points</param>
    /// <param name="fontSize">Text size in points</param>
    /// <param name="fontColor">Text color as #RRGGBB</param>
    /// <param name="fillColor">Fill color as #RRGGBB</param>
    /// <param name="lineColor">Outline color as #RRGGBB</param>
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="connectorType">Connector geometry: Straight, Elbow, or Curved</param>
    /// <param name="beginX">Starting horizontal position in points</param>
    /// <param name="beginY">Starting vertical position in points</param>
    /// <param name="endX">Ending horizontal position in points</param>
    /// <param name="endY">Ending vertical position in points</param>
    /// <param name="name">Optional object name</param>
    /// <param name="lineColor">Line color as #RRGGBB</param>
    /// <param name="lineWeight">Line thickness in points</param>
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="controlType">Worksheet Forms control type; ActiveX/OLE controls are excluded</param>
    /// <param name="name">Optional control name</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points</param>
    /// <param name="height">Height in points</param>
    /// <param name="text">Optional control label</param>
    /// <param name="linkedCell">Cell binding for CheckBox, DropDown, ListBox, OptionButton, ScrollBar, or Spinner</param>
    /// <param name="inputRange">Cell range supplying items to a DropDown or ListBox</param>
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the object</param>
    /// <param name="objectName">Existing object name</param>
    /// <param name="newName">New name for the object; omit to keep its name</param>
    /// <param name="left">New left position in points</param>
    /// <param name="top">New top position in points</param>
    /// <param name="width">New width in points</param>
    /// <param name="height">New height in points</param>
    /// <param name="rotation">Rotation angle in degrees</param>
    /// <param name="text">Replacement text</param>
    /// <param name="fontSize">Text size in points</param>
    /// <param name="fontColor">Text color as #RRGGBB</param>
    /// <param name="fillColor">Fill color as #RRGGBB</param>
    /// <param name="lineColor">Outline color as #RRGGBB</param>
    /// <param name="lineWeight">Line thickness in points</param>
    /// <param name="visible">Show or hide the object; omit to leave unchanged</param>
    /// <param name="locked">Lock the object; effective when worksheet protection is enabled</param>
    /// <param name="placement">Cell anchoring: 1=move and size, 2=move only, 3=free floating</param>
    /// <param name="alternativeText">Accessible description of the object</param>
    /// <param name="linkedCell">Replacement Forms-control cell binding</param>
    /// <param name="inputRange">Replacement item range for a DropDown or ListBox</param>
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the sparkline</param>
    /// <param name="locationRange">Cell or range displaying the sparkline group</param>
    [ServiceAction("get-sparkline")]
    SparklineResult GetSparkline(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string locationRange);

    /// <summary>Adds a line, column, or win/loss sparkline group.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet</param>
    /// <param name="sourceRange">Cell range supplying the sparkline data</param>
    /// <param name="locationRange">Cells where the sparklines will appear</param>
    /// <param name="sparklineType">Sparkline type: Line, Column, or WinLoss</param>
    /// <param name="lineColor">Sparkline color as #RRGGBB</param>
    /// <param name="showMarkers">Display markers on line sparklines</param>
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

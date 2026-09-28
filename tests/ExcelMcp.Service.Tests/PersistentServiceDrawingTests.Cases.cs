using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Drawing;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceDrawingTests
{
    [Fact]
    public void ImageLifecycle_CreateReadUpdateListDelete_RoundTrips()
    {
        var imagePath = CreateTestPng();

        var batch = _fixture.BatchToken;
        var created = _drawingCommands.AddImage(
            batch,
            _sheetName,
            imagePath,
            "ProductImage",
            left: 12,
            top: 18,
            width: 120,
            height: 80);

        Assert.True(created.Success);
        Assert.Equal(DrawingObjectKind.Image, created.DrawingObject.Kind);
        Assert.Equal("ProductImage", created.DrawingObject.Name);

        var read = _drawingCommands.GetObject(batch, _sheetName, "ProductImage");
        Assert.True(read.Success);
        Assert.Equal(12, read.DrawingObject.Left, precision: 1);
        Assert.Equal(120, read.DrawingObject.Width, precision: 1);

        var updated = _drawingCommands.UpdateObject(
            batch,
            _sheetName,
            "ProductImage",
            newName: "RenamedImage",
            left: 42,
            width: 160,
            alternativeText: "Quarterly product image");

        Assert.True(updated.Success);
        Assert.Equal("RenamedImage", updated.DrawingObject.Name);
        Assert.Equal(42, updated.DrawingObject.Left, precision: 1);
        Assert.Equal(160, updated.DrawingObject.Width, precision: 1);
        Assert.Equal("Quarterly product image", updated.DrawingObject.AlternativeText);

        var listed = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.Contains(listed.DrawingObjects, item =>
            item.Name == "RenamedImage" && item.Kind == DrawingObjectKind.Image);

        var deleted = _drawingCommands.DeleteObject(batch, _sheetName, "RenamedImage");
        Assert.True(deleted.Success);
        Assert.DoesNotContain(_drawingCommands.ListObjects(batch, _sheetName).DrawingObjects, item => item.Name == "RenamedImage");
    }

    [Fact]
    public void AutoShapeLifecycle_CreateAndUpdateFormatting_RoundTrips()
    {

        var batch = _fixture.BatchToken;
        var created = _drawingCommands.AddShape(
            batch,
            _sheetName,
            DrawingShapeType.RoundedRectangle,
            "StatusCard",
            left: 25,
            top: 30,
            width: 180,
            height: 75,
            text: "Ready",
            fillColor: "#4472C4",
            lineColor: "#203864",
            lineWeight: 2);

        Assert.True(created.Success);
        Assert.Equal(DrawingObjectKind.AutoShape, created.DrawingObject.Kind);
        Assert.Equal(DrawingShapeType.RoundedRectangle, created.DrawingObject.ShapeType);
        Assert.Equal("Ready", created.DrawingObject.Text);
        Assert.Equal("#4472C4", created.DrawingObject.FillColor);
        Assert.Equal("#203864", created.DrawingObject.LineColor);
        Assert.Equal(2, created.DrawingObject.LineWeight!.Value, precision: 1);

        var updated = _drawingCommands.UpdateObject(
            batch,
            _sheetName,
            "StatusCard",
            text: "Complete",
            fillColor: "#70AD47",
            lineColor: "#385723",
            lineWeight: 3,
            rotation: 5,
            visible: true,
            placement: 2);

        Assert.True(updated.Success);
        Assert.Equal("Complete", updated.DrawingObject.Text);
        Assert.Equal("#70AD47", updated.DrawingObject.FillColor);
        Assert.Equal("#385723", updated.DrawingObject.LineColor);
        Assert.Equal(3, updated.DrawingObject.LineWeight!.Value, precision: 1);
        Assert.Equal(5, updated.DrawingObject.Rotation, precision: 1);
        Assert.Equal(2, updated.DrawingObject.Placement);
    }

    [Fact]
    public void TextBoxConnectorAndSafeFormControls_CreateAndRead_RoundTrip()
    {

        var batch = _fixture.BatchToken;
        var textBox = _drawingCommands.AddTextBox(
            batch,
            _sheetName,
            "Review required",
            "ReviewNote",
            left: 20,
            top: 130,
            width: 200,
            height: 45,
            fontSize: 14,
            fontColor: "#FFFFFF",
            fillColor: "#C00000",
            lineColor: "#7F0000");
        var connector = _drawingCommands.AddConnector(
            batch,
            _sheetName,
            DrawingConnectorType.Elbow,
            beginX: 30,
            beginY: 200,
            endX: 220,
            endY: 250,
            name: "WorkflowConnector",
            lineColor: "#5B9BD5",
            lineWeight: 2.5);
        var checkBox = _drawingCommands.AddFormControl(
            batch,
            _sheetName,
            DrawingFormControlType.CheckBox,
            "ApprovalCheck",
            left: 250,
            top: 25,
            width: 120,
            height: 24,
            text: "Approved",
            linkedCell: $"{_sheetName}!$J$2");
        var dropDown = _drawingCommands.AddFormControl(
            batch,
            _sheetName,
            DrawingFormControlType.DropDown,
            "StatusDropDown",
            left: 250,
            top: 60,
            width: 140,
            height: 24,
            inputRange: $"{_sheetName}!$L$1:$L$3",
            linkedCell: $"{_sheetName}!$J$3");

        Assert.Equal(DrawingObjectKind.TextBox, textBox.DrawingObject.Kind);
        Assert.Equal("Review required", textBox.DrawingObject.Text);
        Assert.Equal(14, textBox.DrawingObject.FontSize!.Value, precision: 1);
        Assert.Equal(DrawingObjectKind.Connector, connector.DrawingObject.Kind);
        Assert.Equal(DrawingConnectorType.Elbow, connector.DrawingObject.ConnectorType);
        Assert.Equal(DrawingObjectKind.FormControl, checkBox.DrawingObject.Kind);
        Assert.Equal(DrawingFormControlType.CheckBox, checkBox.DrawingObject.FormControlType);
        Assert.Equal($"{_sheetName}!$J$2", checkBox.DrawingObject.LinkedCell);
        Assert.Equal(DrawingFormControlType.DropDown, dropDown.DrawingObject.FormControlType);
        Assert.Equal($"{_sheetName}!$L$1:$L$3", dropDown.DrawingObject.InputRange);

        var listed = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.Contains(listed.DrawingObjects, item => item.Name == "ReviewNote");
        Assert.Contains(listed.DrawingObjects, item => item.Name == "WorkflowConnector");
        Assert.Contains(listed.DrawingObjects, item => item.Name == "ApprovalCheck");
        Assert.Contains(listed.DrawingObjects, item => item.Name == "StatusDropDown");
    }

    [Fact]
    public void FormControlsWithoutBindings_ReadExplicitNullProperties()
    {

        var batch = _fixture.BatchToken;
        var controls = new[]
        {
            AddFormControl(batch, DrawingFormControlType.Button, "ActionButton", 20),
            AddFormControl(batch, DrawingFormControlType.GroupBox, "OptionsGroup", 60),
            AddFormControl(batch, DrawingFormControlType.Label, "StatusLabel", 100)
        };

        Assert.All(controls, control =>
        {
            Assert.Null(control.DrawingObject.LinkedCell);
            Assert.Null(control.DrawingObject.InputRange);
        });

        var listed = _drawingCommands.ListObjects(batch, _sheetName).DrawingObjects
            .Where(item => controls.Any(control => control.DrawingObject.Name == item.Name))
            .ToList();
        Assert.Equal(controls.Length, listed.Count);
        Assert.All(
            listed,
            control =>
            {
                Assert.Null(control.LinkedCell);
                Assert.Null(control.InputRange);
            });
    }

    [Fact]
    public void FormControlsWithLinkedCellOnly_ReadLinkedCellAndNullInputRange()
    {

        var batch = _fixture.BatchToken;
        var controls = new[]
        {
            AddFormControl(batch, DrawingFormControlType.CheckBox, "ApprovalCheck", 20, linkedCell: $"{_sheetName}!$J$2"),
            AddFormControl(batch, DrawingFormControlType.OptionButton, "PrimaryOption", 60, linkedCell: $"{_sheetName}!$J$3"),
            AddFormControl(batch, DrawingFormControlType.ScrollBar, "AmountScroll", 100, linkedCell: $"{_sheetName}!$J$4"),
            AddFormControl(batch, DrawingFormControlType.Spinner, "AmountSpinner", 140, linkedCell: $"{_sheetName}!$J$5")
        };

        Assert.All(controls, control =>
        {
            Assert.NotNull(control.DrawingObject.LinkedCell);
            Assert.Null(control.DrawingObject.InputRange);
        });

        var listed = _drawingCommands.ListObjects(batch, _sheetName).DrawingObjects
            .Where(item => controls.Any(control => control.DrawingObject.Name == item.Name))
            .ToList();
        Assert.Equal(controls.Length, listed.Count);
        Assert.All(
            listed,
            control =>
            {
                Assert.NotNull(control.LinkedCell);
                Assert.Null(control.InputRange);
            });
    }

    [Fact]
    public void ListFormControls_ReadLinkedCellAndInputRange()
    {

        var batch = _fixture.BatchToken;
        var controls = new[]
        {
            AddFormControl(
                batch,
                DrawingFormControlType.DropDown,
                "StatusDropDown",
                20,
                linkedCell: $"{_sheetName}!$J$2",
                inputRange: $"{_sheetName}!$L$1:$L$3"),
            AddFormControl(
                batch,
                DrawingFormControlType.ListBox,
                "StatusList",
                60,
                linkedCell: $"{_sheetName}!$J$3",
                inputRange: $"{_sheetName}!$L$1:$L$3")
        };

        Assert.All(controls, control =>
        {
            Assert.NotNull(control.DrawingObject.LinkedCell);
            Assert.NotNull(control.DrawingObject.InputRange);
        });

        var listed = _drawingCommands.ListObjects(batch, _sheetName).DrawingObjects
            .Where(item => controls.Any(control => control.DrawingObject.Name == item.Name))
            .ToList();
        Assert.Equal(controls.Length, listed.Count);
        Assert.All(
            listed,
            control =>
            {
                Assert.NotNull(control.LinkedCell);
                Assert.NotNull(control.InputRange);
            });
    }

    [Fact]
    public void Sparklines_CreateReadUpdateListDelete_RoundTrip()
    {

        var batch = _fixture.BatchToken;
        WriteSparklineData(batch);

        var created = _drawingCommands.AddSparkline(
            batch,
            _sheetName,
            "B2:E2",
            "F2",
            DrawingSparklineType.Line,
            lineColor: "#4472C4",
            showMarkers: true);

        Assert.True(created.Success);
        Assert.Equal("F2", created.Sparkline.LocationRange);
        Assert.Equal("B2:E2", created.Sparkline.SourceRange);
        Assert.Equal(DrawingSparklineType.Line, created.Sparkline.SparklineType);
        Assert.Equal("#4472C4", created.Sparkline.LineColor);
        Assert.True(created.Sparkline.ShowMarkers);

        var read = _drawingCommands.GetSparkline(batch, _sheetName, "F2");
        Assert.True(read.Success);
        Assert.Equal("B2:E2", read.Sparkline.SourceRange);

        var updated = _drawingCommands.UpdateSparkline(
            batch,
            _sheetName,
            "F2",
            sourceRange: "B3:E3",
            sparklineType: DrawingSparklineType.Column,
            lineColor: "#ED7D31",
            showMarkers: false);

        Assert.True(updated.Success);
        Assert.Equal("B3:E3", updated.Sparkline.SourceRange);
        Assert.Equal(DrawingSparklineType.Column, updated.Sparkline.SparklineType);
        Assert.Equal("#ED7D31", updated.Sparkline.LineColor);
        Assert.False(updated.Sparkline.ShowMarkers);

        var listed = _drawingCommands.ListSparklines(batch, _sheetName);
        Assert.Contains(listed.Sparklines, item => item.LocationRange == "F2");

        var deleted = _drawingCommands.DeleteSparkline(batch, _sheetName, "F2");
        Assert.True(deleted.Success);
        Assert.Empty(_drawingCommands.ListSparklines(batch, _sheetName).Sparklines);
    }

    private string CreateTestPng()
    {
        const string onePixelPng =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
        return _fixture.CreateInputFile(
            ".png",
            Convert.FromBase64String(onePixelPng));
    }

    private DrawingObjectResult AddFormControl(
        IExcelBatch batch,
        DrawingFormControlType controlType,
        string name,
        double top,
        string? linkedCell = null,
        string? inputRange = null)
    {
        return _drawingCommands.AddFormControl(
            batch,
            _sheetName,
            controlType,
            name,
            left: 20,
            top: top,
            width: 120,
            height: 24,
            text: SupportsText(controlType) ? name : null,
            linkedCell: linkedCell,
            inputRange: inputRange);
    }

    private static bool SupportsText(DrawingFormControlType controlType)
    {
        return controlType is
            DrawingFormControlType.Button or
            DrawingFormControlType.CheckBox or
            DrawingFormControlType.GroupBox or
            DrawingFormControlType.Label or
            DrawingFormControlType.OptionButton;
    }

    private void WriteSparklineData(IExcelBatch batch)
    {
        _rangeCommands.SetValues(
            batch,
            _sheetName,
            "B2:E3",
            [[1, 3, 2, 5], [5, 2, 4, 1]]);
    }
}

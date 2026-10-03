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
        var actual = _drawingCommands.GetObject(batch, _sheetName, "StatusCard");
        Assert.True(actual.Success, actual.ErrorMessage);
        Assert.Equal("Complete", actual.DrawingObject.Text);
        Assert.Equal("#70AD47", actual.DrawingObject.FillColor);
        Assert.Equal("#385723", actual.DrawingObject.LineColor);
        Assert.Equal(3, actual.DrawingObject.LineWeight!.Value, precision: 1);
        Assert.Equal(5, actual.DrawingObject.Rotation, precision: 1);
        Assert.Equal(2, actual.DrawingObject.Placement);
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

        var read = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(read.Success, read.ErrorMessage);
        var listed = read.DrawingObjects
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

        var read = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(read.Success, read.ErrorMessage);
        var listed = read.DrawingObjects
            .Where(item => controls.Any(control => control.DrawingObject.Name == item.Name))
            .ToList();
        Assert.Equal(controls.Length, listed.Count);
        Assert.All(
            listed,
            control =>
            {
                var index = Array.FindIndex(controls, original => original.DrawingObject.Name == control.Name);
                Assert.Equal($"{_sheetName}!$J${index + 2}", control.LinkedCell);
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

        var read = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(read.Success, read.ErrorMessage);
        var listed = read.DrawingObjects
            .Where(item => controls.Any(control => control.DrawingObject.Name == item.Name))
            .ToList();
        Assert.Equal(controls.Length, listed.Count);
        Assert.All(
            listed,
            control =>
            {
                var index = Array.FindIndex(controls, original => original.DrawingObject.Name == control.Name);
                Assert.Equal($"{_sheetName}!$J${index + 2}", control.LinkedCell);
                Assert.Equal($"{_sheetName}!$L$1:$L$3", control.InputRange);
            });
    }

    [Theory]
    [InlineData(DrawingFormControlType.CheckBox)]
    [InlineData(DrawingFormControlType.OptionButton)]
    [InlineData(DrawingFormControlType.ScrollBar)]
    [InlineData(DrawingFormControlType.Spinner)]
    [InlineData(DrawingFormControlType.DropDown)]
    [InlineData(DrawingFormControlType.ListBox)]
    public void UpdateObject_SupportedFormBindings_ChangesActualBindings(DrawingFormControlType controlType)
    {
        var batch = _fixture.BatchToken;
        var supportsInputRange = controlType is DrawingFormControlType.DropDown or DrawingFormControlType.ListBox;
        Assert.True(_rangeCommands.SetValues(batch, _sheetName, "L1:L6",
            [["First"], ["Second"], ["Third"], ["Fourth"], ["Fifth"], ["Sixth"]]).Success);
        var created = AddFormControl(batch, controlType, "BoundControl", 20,
            linkedCell: $"{_sheetName}!$J$2",
            inputRange: supportsInputRange ? $"{_sheetName}!$L$1:$L$3" : null);
        Assert.True(created.Success, created.ErrorMessage);
        var before = _drawingCommands.GetObject(batch, _sheetName, "BoundControl");
        Assert.True(before.Success, before.ErrorMessage);
        Assert.Equal($"{_sheetName}!$J$2", before.DrawingObject.LinkedCell);
        Assert.Equal(supportsInputRange ? $"{_sheetName}!$L$1:$L$3" : null, before.DrawingObject.InputRange);

        var updated = _drawingCommands.UpdateObject(batch, _sheetName, "BoundControl",
            linkedCell: $"{_sheetName}!$J$9",
            inputRange: supportsInputRange ? $"{_sheetName}!$L$4:$L$6" : null);
        Assert.True(updated.Success, updated.ErrorMessage);
        var after = _drawingCommands.GetObject(batch, _sheetName, "BoundControl");
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal($"{_sheetName}!$J$9", after.DrawingObject.LinkedCell);
        Assert.Equal(supportsInputRange ? $"{_sheetName}!$L$4:$L$6" : null, after.DrawingObject.InputRange);
        Assert.Equal(before.DrawingObject.Name, after.DrawingObject.Name);
        Assert.Equal(before.DrawingObject.Left, after.DrawingObject.Left);
        Assert.Equal(before.DrawingObject.Top, after.DrawingObject.Top);
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
        var actual = _drawingCommands.GetSparkline(batch, _sheetName, "F2");
        Assert.True(actual.Success, actual.ErrorMessage);
        Assert.Equal("B3:E3", actual.Sparkline.SourceRange);
        Assert.Equal(DrawingSparklineType.Column, actual.Sparkline.SparklineType);
        Assert.Equal("#ED7D31", actual.Sparkline.LineColor);
        Assert.False(actual.Sparkline.ShowMarkers);

        var listed = _drawingCommands.ListSparklines(batch, _sheetName);
        Assert.Contains(listed.Sparklines, item => item.LocationRange == "F2");

        var deleted = _drawingCommands.DeleteSparkline(batch, _sheetName, "F2");
        Assert.True(deleted.Success);
        Assert.Empty(_drawingCommands.ListSparklines(batch, _sheetName).Sparklines);
    }

    [Theory]
    [InlineData("font")]
    [InlineData("fill")]
    [InlineData("line")]
    public void UpdateObject_InvalidColor_PreservesNameGeometryAndText(string target)
    {
        var batch = _fixture.BatchToken;
        var created = _drawingCommands.AddShape(batch, _sheetName, DrawingShapeType.Rectangle,
            "RetainedShape", left: 10, top: 20, text: "Before", fillColor: "#4472C4");
        Assert.True(created.Success, created.ErrorMessage);

        var exception = Assert.Throws<ArgumentException>(() =>
            _drawingCommands.UpdateObject(batch, _sheetName, "RetainedShape",
                newName: "RejectedRename", left: 99, text: "After",
                fontColor: target == "font" ? "invalid" : null,
                fillColor: target == "fill" ? "invalid" : null,
                lineColor: target == "line" ? "invalid" : null));

        Assert.Contains("Invalid color", exception.Message, StringComparison.Ordinal);
        AssertRetainedShape(batch);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void UpdateObject_BindingsOnNonControl_PreservesObject(bool linkedCell)
    {
        var batch = _fixture.BatchToken;
        var created = _drawingCommands.AddShape(batch, _sheetName, DrawingShapeType.Rectangle,
            "RetainedShape", left: 10, top: 20, text: "Before", fillColor: "#4472C4");
        Assert.True(created.Success, created.ErrorMessage);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _drawingCommands.UpdateObject(batch, _sheetName, "RetainedShape",
                newName: "RejectedRename", left: 99, text: "After",
                linkedCell: linkedCell ? $"{_sheetName}!$A$1" : null,
                inputRange: linkedCell ? null : $"{_sheetName}!$A$1:$A$3"));

        Assert.Contains("apply only to worksheet Forms controls", exception.Message, StringComparison.Ordinal);
        AssertRetainedShape(batch);
    }

    private void AssertRetainedShape(IExcelBatch batch)
    {
        var listed = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal("RetainedShape", Assert.Single(listed.DrawingObjects).Name);
        var actual = _drawingCommands.GetObject(batch, _sheetName, "RetainedShape");
        Assert.True(actual.Success, actual.ErrorMessage);
        Assert.Equal(10, actual.DrawingObject.Left, precision: 1);
        Assert.Equal(20, actual.DrawingObject.Top, precision: 1);
        Assert.Equal("Before", actual.DrawingObject.Text);
        Assert.Equal("#4472C4", actual.DrawingObject.FillColor);
    }

    [Theory]
    [InlineData(DrawingFormControlType.Button, true, false)]
    [InlineData(DrawingFormControlType.GroupBox, true, false)]
    [InlineData(DrawingFormControlType.Label, true, false)]
    [InlineData(DrawingFormControlType.Button, false, false)]
    [InlineData(DrawingFormControlType.GroupBox, false, false)]
    [InlineData(DrawingFormControlType.Label, false, false)]
    [InlineData(DrawingFormControlType.CheckBox, false, false)]
    [InlineData(DrawingFormControlType.OptionButton, false, false)]
    [InlineData(DrawingFormControlType.ScrollBar, false, false)]
    [InlineData(DrawingFormControlType.Spinner, false, false)]
    [InlineData(DrawingFormControlType.Button, true, true)]
    [InlineData(DrawingFormControlType.GroupBox, true, true)]
    [InlineData(DrawingFormControlType.Label, true, true)]
    [InlineData(DrawingFormControlType.Button, false, true)]
    [InlineData(DrawingFormControlType.GroupBox, false, true)]
    [InlineData(DrawingFormControlType.Label, false, true)]
    [InlineData(DrawingFormControlType.CheckBox, false, true)]
    [InlineData(DrawingFormControlType.OptionButton, false, true)]
    [InlineData(DrawingFormControlType.ScrollBar, false, true)]
    [InlineData(DrawingFormControlType.Spinner, false, true)]
    public void FormControl_UnsupportedBinding_PreservesObjectsAndCells(
        DrawingFormControlType controlType, bool linkedCell, bool update)
    {
        var batch = _fixture.BatchToken;
        Assert.True(_rangeCommands.SetValues(batch, _sheetName, "A1:A3",
            [["First"], ["Second"], ["Third"]]).Success);
        DrawingObjectInfo? original = null;
        if (update)
        {
            var created = AddFormControl(batch, controlType, "RetainedControl", 20);
            Assert.True(created.Success, created.ErrorMessage);
            original = created.DrawingObject;
        }

        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            if (update)
                _drawingCommands.UpdateObject(batch, _sheetName, "RetainedControl",
                    newName: "RejectedRename", left: 99, top: 88,
                    linkedCell: linkedCell ? $"{_sheetName}!$A$1" : null,
                    inputRange: linkedCell ? null : $"{_sheetName}!$A$1:$A$3");
            else
                _drawingCommands.AddFormControl(batch, _sheetName, controlType, "RejectedControl",
                    linkedCell: linkedCell ? $"{_sheetName}!$A$1" : null,
                    inputRange: linkedCell ? null : $"{_sheetName}!$A$1:$A$3");
        });

        var listed = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        if (update)
        {
            var retained = Assert.Single(listed.DrawingObjects);
            Assert.NotNull(original);
            Assert.Equal(original.Name, retained.Name);
            Assert.Equal(original.Left, retained.Left);
            Assert.Equal(original.Top, retained.Top);
            Assert.Equal(original.Width, retained.Width);
            Assert.Equal(original.Height, retained.Height);
            Assert.Equal(original.Text, retained.Text);
            Assert.Equal(original.LinkedCell, retained.LinkedCell);
            Assert.Equal(original.InputRange, retained.InputRange);
        }
        else
        {
            Assert.Empty(listed.DrawingObjects);
        }
        var cells = _rangeCommands.GetValues(batch, _sheetName, "A1:A3");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["First", "Second", "Third"], cells.Values.Select(row => row[0]?.ToString()));
        Assert.Contains("not supported", exception.Message, StringComparison.Ordinal);
        Assert.Contains(linkedCell ? "linkedCell" : "inputRange", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("shape-fill")]
    [InlineData("shape-line")]
    [InlineData("text-font")]
    [InlineData("text-fill")]
    [InlineData("text-line")]
    [InlineData("connector")]
    public void AddObject_InvalidColor_DoesNotLeaveAnObject(string target)
    {
        var batch = _fixture.BatchToken;

        var exception = Assert.Throws<ArgumentException>(() =>
        {
            if (target.StartsWith("shape", StringComparison.Ordinal))
                _drawingCommands.AddShape(batch, _sheetName, name: "Rejected",
                    fillColor: target == "shape-fill" ? "invalid" : null,
                    lineColor: target == "shape-line" ? "invalid" : null);
            else if (target.StartsWith("text", StringComparison.Ordinal))
                _drawingCommands.AddTextBox(batch, _sheetName, "Rejected", name: "Rejected",
                    fontColor: target == "text-font" ? "invalid" : null,
                    fillColor: target == "text-fill" ? "invalid" : null,
                    lineColor: target == "text-line" ? "invalid" : null);
            else
                _drawingCommands.AddConnector(batch, _sheetName, name: "Rejected", lineColor: "invalid");
        });

        Assert.Contains("Invalid color", exception.Message, StringComparison.Ordinal);
        var listed = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Empty(listed.DrawingObjects);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Sparkline_InvalidColor_PreservesExistingGroups(bool update)
    {
        var batch = _fixture.BatchToken;
        WriteSparklineData(batch);
        var created = _drawingCommands.AddSparkline(batch, _sheetName, "B2:E2", "F2",
            lineColor: "#4472C4", showMarkers: true);
        Assert.True(created.Success, created.ErrorMessage);

        var exception = Assert.Throws<ArgumentException>(() =>
        {
            if (update)
                _drawingCommands.UpdateSparkline(batch, _sheetName, "F2",
                    sourceRange: "B3:E3", sparklineType: DrawingSparklineType.Column,
                    lineColor: "invalid", showMarkers: false);
            else
                _drawingCommands.AddSparkline(batch, _sheetName, "B3:E3", "F3", lineColor: "invalid");
        });

        Assert.Contains("Invalid color", exception.Message, StringComparison.Ordinal);
        var listed = _drawingCommands.ListSparklines(batch, _sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        var retained = Assert.Single(listed.Sparklines);
        Assert.Equal("F2", retained.LocationRange);
        Assert.Equal("B2:E2", retained.SourceRange);
        Assert.Equal(DrawingSparklineType.Line, retained.SparklineType);
        Assert.Equal("#4472C4", retained.LineColor);
        Assert.True(retained.ShowMarkers);
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
        var result = _drawingCommands.AddFormControl(
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
        Assert.True(result.Success, result.ErrorMessage);
        return result;
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

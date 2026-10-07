using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceDrawingTests
{
    // Excel accepts drawing names up to 254 characters.
    private static readonly string NameExcelRejects = new('N', 300);

    [Theory]
    [InlineData("image")]
    [InlineData("shape")]
    [InlineData("text box")]
    [InlineData("connector")]
    [InlineData("form control")]
    public void AddObject_OverlongName_IsRejectedBeforeAddingObject(string kind)
    {
        var batch = _fixture.BatchToken;

        var exception = Assert.Throws<ArgumentException>(() =>
        {
            _ = kind switch
            {
                "image" => _drawingCommands.AddImage(batch, _sheetName, CreateTestPng(), NameExcelRejects),
                "shape" => _drawingCommands.AddShape(batch, _sheetName, name: NameExcelRejects),
                "text box" => _drawingCommands.AddTextBox(batch, _sheetName, "Kept", name: NameExcelRejects),
                "connector" => _drawingCommands.AddConnector(batch, _sheetName, name: NameExcelRejects),
                _ => _drawingCommands.AddFormControl(batch, _sheetName, name: NameExcelRejects),
            };
        });

        var listed = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Empty(listed.DrawingObjects);
        Assert.Contains("254", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("connector")]
    public void AddObject_NegativeLineWeight_DoesNotLeaveAnObject(string kind)
    {
        var batch = _fixture.BatchToken;

        var exception = Assert.Throws<ArgumentException>(() =>
        {
            if (kind == "shape")
                _drawingCommands.AddShape(batch, _sheetName, name: "Rejected", lineWeight: -1);
            else
                _drawingCommands.AddConnector(batch, _sheetName, name: "Rejected", lineWeight: -1);
        });

        Assert.Contains("lineWeight", exception.Message, StringComparison.Ordinal);
        var listed = _drawingCommands.ListObjects(batch, _sheetName);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Empty(listed.DrawingObjects);
    }
}
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Drawing;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceDrawingTests
{
    [Theory]
    [InlineData(DrawingConnectorType.Straight)]
    [InlineData(DrawingConnectorType.Elbow)]
    [InlineData(DrawingConnectorType.Curved)]
    public void Connector_ReadbackSkipsUnsupportedTextFrameAndPreservesGeometry(DrawingConnectorType type)
    {
        const string name = "NativeConnector";
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.TextFrame? frame = null;
            Excel.Characters? characters = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, _sheetName);
                Assert.NotNull(sheet);
                shapes = sheet.Shapes;
                shape = shapes.AddConnector((Office.MsoConnectorType)type, 30f, 40f, 130f, 90f);
                shape.Name = name;
                var error = Assert.Throws<COMException>(() =>
                {
                    frame = shape.TextFrame;
                    characters = frame.Characters();
                });
                Assert.Equal(type == DrawingConnectorType.Straight
                    ? ComErrors.MemberNotFound : unchecked((int)0x800A03EC), error.HResult);
                Assert.Equal(1, shapes.Count);
                Assert.Equal(name, shape.Name);
            }
            finally
            {
                ComUtilities.Release(ref characters);
                ComUtilities.Release(ref frame);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
        var info = RequireSuccess(_drawingCommands.GetObject(_fixture.BatchToken, _sheetName, name)).DrawingObject;
        Assert.Equal(DrawingObjectKind.Connector, info.Kind);
        Assert.Equal(type, info.ConnectorType);
        Assert.Equal(30, info.Left, precision: 1);
        Assert.Equal(40, info.Top, precision: 1);
        Assert.Equal(100, info.Width, precision: 1);
        Assert.Equal(50, info.Height, precision: 1);
        Assert.Null(info.Text);
        Assert.Null(info.FontSize);
        Assert.Null(info.FontColor);
    }
}

using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void AddShape_InsertsShapeIntoWorksheetAndCanBeCounted()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        SeedObjectGuardCells(sheetName);
        var before = CaptureObjectGuardCells(sheetName);

        RequireSuccess(_sheetCommands.AddShape(batch, sheetName, "B2"));

        var countResult = RequireSuccess(_sheetCommands.GetShapeCount(batch, sheetName));
        Assert.Equal(1, countResult.ShapeCount);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.Range? anchor = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                shapes = sheet.Shapes;
                Assert.Equal(1, shapes.Count);
                shape = (Excel.Shape)shapes.Item(1);
                // PIA gap: AutoShapeType returns an unavailable Office-core enum.
                Assert.Equal(1, Convert.ToInt32(((dynamic)shape).AutoShapeType,
                    System.Globalization.CultureInfo.InvariantCulture));
                anchor = sheet.Range["B2"];
                Assert.Equal(anchor.Left, shape.Left);
                Assert.Equal(anchor.Top, shape.Top);
                Assert.Equal(144d, shape.Width);
                Assert.Equal(72d, shape.Height);
            }
            finally
            {
                ComUtilities.Release(ref anchor);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
        Assert.Equal(before, CaptureObjectGuardCells(sheetName));
    }

    [Fact]
    public void AddShape_InvalidCell_PreservesExistingWorksheetAndShapes()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        SeedObjectGuardCells(sheetName);
        var before = CaptureObjectGuardCells(sheetName);

        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.AddShape(batch, sheetName, "not-a-cell"));

        Assert.Contains("sheet.add-shape failed", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, CaptureObjectGuardCells(sheetName));
        Assert.Equal(0, RequireSuccess(_sheetCommands.GetShapeCount(batch, sheetName)).ShapeCount);
    }

    [Fact]
    public void AddImage_InsertsPictureIntoWorksheetAndCanBeCounted()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        SeedObjectGuardCells(sheetName);
        var before = CaptureObjectGuardCells(sheetName);
        var imagePath = _fixture.CreateInputFile(
            ".png",
            Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAIAAeIhvAAAAAElFTkSuQmCC"));

        RequireSuccess(_sheetCommands.AddImage(
            batch,
            sheetName,
            imagePath,
            "B2"));

        var countResult = RequireSuccess(_sheetCommands.GetImageCount(batch, sheetName));
        Assert.Equal(1, countResult.ImageCount);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Pictures? pictures = null;
            Excel.Picture? picture = null;
            Excel.Range? anchor = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                pictures = (Excel.Pictures)sheet.Pictures(Type.Missing);
                Assert.Equal(1, pictures.Count);
                picture = (Excel.Picture)pictures.Item(1);
                anchor = sheet.Range["B2"];
                Assert.Equal(anchor.Left, picture.Left);
                Assert.Equal(anchor.Top, picture.Top);
                Assert.True(picture.Width > 0);
                Assert.True(picture.Height > 0);
            }
            finally
            {
                ComUtilities.Release(ref anchor);
                ComUtilities.Release(ref picture);
                ComUtilities.Release(ref pictures);
                ComUtilities.Release(ref sheet);
            }
        });
        Assert.Equal(before, CaptureObjectGuardCells(sheetName));
    }

    [Fact]
    public void AddImage_MissingFile_PreservesExistingWorksheetAndImages()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        SeedObjectGuardCells(sheetName);
        var before = CaptureObjectGuardCells(sheetName);
        var missingPath = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.png");
        Assert.False(File.Exists(missingPath));

        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.AddImage(batch, sheetName, missingPath, "B2"));

        Assert.Contains("sheet.add-image failed", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, CaptureObjectGuardCells(sheetName));
        Assert.Equal(0, RequireSuccess(_sheetCommands.GetImageCount(batch, sheetName)).ImageCount);
    }

    [Fact]
    public void SetProtection_ProtectsAndUnprotectsSheet()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        SeedObjectGuardCells(sheetName);
        var before = CaptureObjectGuardCells(sheetName);

        RequireSuccess(_sheetCommands.SetProtection(batch, sheetName, true, "test-password"));
        var protectedState = RequireSuccess(_sheetCommands.GetProtection(batch, sheetName));
        Assert.True(protectedState.IsProtected);
        AssertNativeProtection(sheetName, true);
        Assert.Equal(before, CaptureObjectGuardCells(sheetName));

        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.SetProtection(batch, sheetName, false, "wrong-password"));
        Assert.Contains("sheet.set-protection failed", error.Message, StringComparison.Ordinal);
        Assert.True(RequireSuccess(_sheetCommands.GetProtection(batch, sheetName)).IsProtected);
        AssertNativeProtection(sheetName, true);
        Assert.Equal(before, CaptureObjectGuardCells(sheetName));

        RequireSuccess(_sheetCommands.SetProtection(batch, sheetName, false, "test-password"));
        Assert.False(RequireSuccess(_sheetCommands.GetProtection(batch, sheetName)).IsProtected);
        AssertNativeProtection(sheetName, false);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A3", [["updated after unprotect"]]));
        Assert.Equal("updated after unprotect", RequireSuccess(_commands.GetValues(batch, sheetName, "A3")).Values[0][0]);
    }

    private void SeedObjectGuardCells(string sheetName)
    {
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B2",
            [["retained object guard", 17], ["neighbor", 23]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "C2", [["=B2*3"]]));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheetName, "B1:B2", "0.00"));
    }

    private string CaptureObjectGuardCells(string sheetName)
    {
        var formulas = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "A1:C2"));
        var formats = RequireSuccess(_commands.GetNumberFormats(_fixture.BatchToken, sheetName, "A1:C2"));
        return System.Text.Json.JsonSerializer.Serialize(new
        {
            formulas.Values,
            formulas.Formulas,
            formats.Formats
        });
    }

    private void AssertNativeProtection(string sheetName, bool expected)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                Assert.Equal(expected, sheet.ProtectContents);
            }
            finally { ComUtilities.Release(ref sheet); }
        });
    }
}

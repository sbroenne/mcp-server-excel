using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void AddShape_InsertsShapeIntoWorksheetAndCanBeCounted()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var addResult = _sheetCommands.AddShape(batch, sheetName, "A1");
        Assert.True(
            addResult.Success,
            $"Expected shape add to succeed but got error: {addResult.ErrorMessage}");

        var countResult = _sheetCommands.GetShapeCount(batch, sheetName);
        Assert.True(countResult.Success);
        Assert.True(countResult.ShapeCount > 0);
    }

    [Fact]
    public void AddImage_InsertsPictureIntoWorksheetAndCanBeCounted()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var imagePath = _fixture.CreateInputFile(
            ".png",
            Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAIAAeIhvAAAAAElFTkSuQmCC"));

        var addResult = _sheetCommands.AddImage(
            batch,
            sheetName,
            imagePath,
            "A1");
        Assert.True(
            addResult.Success,
            $"Expected image add to succeed but got error: {addResult.ErrorMessage}");

        var countResult = _sheetCommands.GetImageCount(batch, sheetName);
        Assert.True(countResult.Success);
        Assert.True(countResult.ImageCount > 0);
    }

    [Fact]
    public void SetProtection_ProtectsAndUnprotectsSheet()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var protectResult = _sheetCommands.SetProtection(batch, sheetName, true);
        Assert.True(
            protectResult.Success,
            $"Expected protect to succeed but got error: {protectResult.ErrorMessage}");
        Assert.True(_sheetCommands.GetProtection(batch, sheetName).IsProtected);

        var unprotectResult = _sheetCommands.SetProtection(batch, sheetName, false);
        Assert.True(
            unprotectResult.Success,
            $"Expected unprotect to succeed but got error: {unprotectResult.ErrorMessage}");
        Assert.False(_sheetCommands.GetProtection(batch, sheetName).IsProtected);
    }

    [Fact]
    public void GetPageSetup_AutomaticScaling_ReturnsNullFitValues()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = _sheetCommands.GetPageSetup(batch, sheetName);

        Assert.True(
            result.Success,
            $"Expected page setup read to succeed but got error: {result.ErrorMessage}");
        Assert.Null(result.FitToPagesWide);
        Assert.Null(result.FitToPagesTall);
    }
}

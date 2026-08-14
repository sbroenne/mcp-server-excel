using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for worksheet image operations.
/// </summary>
public partial class SheetCommandsTests
{
    [Fact]
    public void AddImage_InsertsPictureIntoWorksheetAndCanBeCounted()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Images_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        var imagePath = Path.Combine(_fixture.TempDir, "sample.png");
        System.IO.File.WriteAllBytes(imagePath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAIAAeIhvAAAAAElFTkSuQmCC"));

        var addResult = _sheetCommands.AddImage(batch, sheetName, imagePath, "A1");
        Assert.True(addResult.Success, $"Expected image add to succeed but got error: {addResult.ErrorMessage}");

        var countResult = _sheetCommands.GetImageCount(batch, sheetName);
        Assert.True(countResult.Success);
        Assert.True(countResult.ImageCount > 0);
    }
}

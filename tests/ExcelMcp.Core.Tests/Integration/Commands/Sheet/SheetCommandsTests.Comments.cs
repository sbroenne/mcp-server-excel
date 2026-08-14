using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for worksheet comments.
/// </summary>
public partial class SheetCommandsTests
{
    [Fact]
    public void SetComment_RoundsTripThroughSheetAndCanBeCleared()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Comments_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        var initialComment = _sheetCommands.GetComment(batch, sheetName, "A1");
        Assert.True(initialComment.Success);
        Assert.False(initialComment.HasComment);

        var setResult = _sheetCommands.SetComment(batch, sheetName, "A1", "Quarterly update");
        Assert.True(setResult.Success, $"Expected comment set to succeed but got error: {setResult.ErrorMessage}");

        var readResult = _sheetCommands.GetComment(batch, sheetName, "A1");
        Assert.True(readResult.Success);
        Assert.True(readResult.HasComment);
        Assert.Equal("Quarterly update", readResult.Text);

        var clearResult = _sheetCommands.ClearComment(batch, sheetName, "A1");
        Assert.True(clearResult.Success, $"Expected comment clear to succeed but got error: {clearResult.ErrorMessage}");

        var clearedResult = _sheetCommands.GetComment(batch, sheetName, "A1");
        Assert.True(clearedResult.Success);
        Assert.False(clearedResult.HasComment);
        Assert.Null(clearedResult.Text);
    }
}

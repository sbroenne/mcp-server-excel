using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for worksheet comments.
/// </summary>
public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void SetComment_RoundsTripThroughSheetAndCanBeCleared()
    {
        var batch = _fixture.BatchToken;
        var sheetName = $"Comments_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        var initialComment = _sheetCommands.GetComment(batch, sheetName, "A1");
        RequireSuccess(initialComment);
        Assert.False(initialComment.HasComment);

        var setResult = _sheetCommands.SetComment(batch, sheetName, "A1", "Quarterly update");
        RequireSuccess(setResult);

        var readResult = _sheetCommands.GetComment(batch, sheetName, "A1");
        RequireSuccess(readResult);
        Assert.True(readResult.HasComment);
        Assert.Equal("Quarterly update", readResult.Text);

        var clearResult = _sheetCommands.ClearComment(batch, sheetName, "A1");
        RequireSuccess(clearResult);

        var clearedResult = _sheetCommands.GetComment(batch, sheetName, "A1");
        RequireSuccess(clearedResult);
        Assert.False(clearedResult.HasComment);
        Assert.Null(clearedResult.Text);
    }
}

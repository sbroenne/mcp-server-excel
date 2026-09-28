using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void Move_WithBeforeSheet_RepositionsSheet()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "MoveMe");
        _fixture.CreateNamedTestSheet(batch, "Target");

        _sheetCommands.Move(batch, "MoveMe", beforeSheet: "Sheet1");

        var sheets = _sheetCommands.List(batch).Worksheets.ToList();
        var movedIndex = sheets.FindIndex(sheet => sheet.Name == "MoveMe");
        var sheet1Index = sheets.FindIndex(sheet => sheet.Name == "Sheet1");
        Assert.True(
            movedIndex < sheet1Index,
            $"Expected MoveMe (index {movedIndex}) before Sheet1 (index {sheet1Index}).");
    }

    [Fact]
    public void Move_WithAfterSheet_RepositionsSheet()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "MoveMe");
        _fixture.CreateNamedTestSheet(batch, "Target");

        _sheetCommands.Move(batch, "MoveMe", afterSheet: "Target");

        var sheets = _sheetCommands.List(batch).Worksheets.ToList();
        var movedIndex = sheets.FindIndex(sheet => sheet.Name == "MoveMe");
        var targetIndex = sheets.FindIndex(sheet => sheet.Name == "Target");
        Assert.True(
            movedIndex > targetIndex,
            $"Expected MoveMe (index {movedIndex}) after Target (index {targetIndex}).");
    }

    [Fact]
    public void Move_NoPositionSpecified_MovesToEnd()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Sheet2");
        _fixture.CreateNamedTestSheet(batch, "Sheet3");

        _sheetCommands.Move(batch, "Sheet1");

        var result = _sheetCommands.List(batch);
        Assert.True(result.Success);
        Assert.Equal("Sheet1", result.Worksheets[^1].Name);
    }

    [Fact]
    public void Move_BothBeforeAndAfter_ThrowsException()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Sheet2");
        _fixture.CreateNamedTestSheet(batch, "Sheet3");

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.Move(
                batch,
                "Sheet1",
                beforeSheet: "Sheet2",
                afterSheet: "Sheet3"));

        Assert.Contains(
            "both beforeSheet and afterSheet",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Move_NonExistentSheet_ThrowsException()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.Move(
                _fixture.BatchToken,
                "NonExistent",
                afterSheet: "Sheet1"));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Move_NonExistentTargetSheet_ThrowsException()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Sheet2");

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.Move(batch, "Sheet2", beforeSheet: "NonExistent"));

        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}

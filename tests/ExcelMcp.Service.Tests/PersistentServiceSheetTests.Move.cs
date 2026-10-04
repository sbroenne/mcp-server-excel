using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void Move_WithBeforeSheet_RepositionsSheet() => AssertMove("before");

    [Fact]
    public void Move_WithAfterSheet_RepositionsSheet() => AssertMove("after");

    [Fact]
    public void Move_NoPositionSpecified_MovesToEnd() => AssertMove("end");

    private void AssertMove(string position)
    {
        var source = CreateSeededSheet("Move");
        var target = CreateSeededSheet("Target");
        var before = ReadCurrentSheetNames();
        var sourceState = CaptureSheet(source);
        var targetState = CaptureSheet(target);
        var expected = before.Where(name => name != source).ToList();
        var index = position == "end" ? expected.Count : expected.IndexOf(target) + (position == "after" ? 1 : 0);
        Assert.InRange(index, 0, expected.Count);
        expected.Insert(index, source);
        RequireSuccess(_sheetCommands.Move(_fixture.BatchToken, source,
            beforeSheet: position == "before" ? target : null,
            afterSheet: position == "after" ? target : null));
        Assert.Equal(expected, ReadCurrentSheetNames());
        Assert.Equal(sourceState, CaptureSheet(source));
        Assert.Equal(targetState, CaptureSheet(target));
    }

    [Fact]
    public void Move_BothBeforeAndAfter_ThrowsException()
    {
        var source = CreateSeededSheet("Move");
        var target = CreateSeededSheet("Target");
        var names = ReadCurrentSheetNames();
        var sourceState = CaptureSheet(source);
        var targetState = CaptureSheet(target);
        var error = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.Move(_fixture.BatchToken, source, beforeSheet: target, afterSheet: target));
        Assert.Contains("both beforeSheet and afterSheet", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(names, ReadCurrentSheetNames());
        Assert.Equal(sourceState, CaptureSheet(source));
        Assert.Equal(targetState, CaptureSheet(target));
    }

    [Fact]
    public void Move_NonExistentSheet_ThrowsException() => AssertMissingMove(missingSource: true);

    [Fact]
    public void Move_NonExistentTargetSheet_ThrowsException() => AssertMissingMove(missingSource: false);

    private void AssertMissingMove(bool missingSource)
    {
        var source = CreateSeededSheet("Move");
        var target = CreateSeededSheet("Target");
        var names = ReadCurrentSheetNames();
        var sourceState = CaptureSheet(source);
        var targetState = CaptureSheet(target);
        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.Move(_fixture.BatchToken, missingSource ? "MissingSheet" : source,
                beforeSheet: missingSource ? target : "MissingSheet"));
        Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(names, ReadCurrentSheetNames());
        Assert.Equal(sourceState, CaptureSheet(source));
        Assert.Equal(targetState, CaptureSheet(target));
    }
}

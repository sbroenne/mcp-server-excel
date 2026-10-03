using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceWindowTests
{
    [Theory]
    [InlineData("left-half")]
    [InlineData("right-half")]
    [InlineData("top-half")]
    [InlineData("bottom-half")]
    [InlineData("center")]
    [InlineData("full-screen")]
    public void Arrange_ValidPresets_Succeed(string preset)
    {
        var batch = _fixture.BatchToken;
        Assert.True(_commands.SetPosition(batch, left: 10, top: 10, width: 300, height: 250).Success);
        var workArea = GetExcelMonitorWorkArea();

        var result = _commands.Arrange(batch, preset);

        Assert.True(result.Success, $"Arrange '{preset}' failed: {result.ErrorMessage}");
        Assert.Equal("arrange", result.Action);
        Assert.Contains(preset, result.Message, StringComparison.OrdinalIgnoreCase);
        var actual = _commands.GetInfo(batch);
        Assert.True(actual.Success, actual.ErrorMessage);
        Assert.True(actual.IsVisible);
        Assert.Equal(preset == "full-screen" ? "maximized" : "normal", actual.WindowState);
        if (preset != "full-screen")
        {
            var expected = GetExpectedBounds(preset, workArea);
            AssertClose(expected.Left, actual.Left);
            AssertClose(expected.Top, actual.Top);
            AssertClose(expected.Width, actual.Width);
            AssertClose(expected.Height, actual.Height);
        }
    }

    [Theory]
    [InlineData("normal", true)]
    [InlineData("normal", false)]
    [InlineData("minimized", true)]
    [InlineData("minimized", false)]
    [InlineData("maximized", true)]
    [InlineData("maximized", false)]
    public void Arrange_InvalidPreset_Throws(string state, bool visible)
    {
        var batch = _fixture.BatchToken;
        Assert.True(_commands.SetState(batch, state).Success);
        if (!visible)
        {
            Assert.True(_commands.Hide(batch).Success);
        }
        var before = _commands.GetInfo(batch);
        Assert.True(before.Success, before.ErrorMessage);
        if (visible)
        {
            Assert.Equal(state, before.WindowState);
        }
        Assert.Equal(visible, before.IsVisible);
        var nativeBefore = _fixture.ExecuteRawVerification((context, _) =>
            (context.App.Visible, context.App.WindowState, context.App.Left,
                context.App.Top, context.App.Width, context.App.Height));

        var exception = Assert.Throws<ArgumentException>(() => _commands.Arrange(batch, "invalid-preset"));

        Assert.Contains("Unknown arrange preset", exception.Message, StringComparison.Ordinal);
        var after = _commands.GetInfo(batch);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(before.IsVisible, after.IsVisible);
        Assert.Equal(before.WindowState, after.WindowState);
        Assert.Equal(before.Left, after.Left);
        Assert.Equal(before.Top, after.Top);
        Assert.Equal(before.Width, after.Width);
        Assert.Equal(before.Height, after.Height);
        var nativeAfter = _fixture.ExecuteRawVerification((context, _) =>
            (context.App.Visible, context.App.WindowState, context.App.Left,
                context.App.Top, context.App.Width, context.App.Height));
        Assert.Equal(nativeBefore, nativeAfter);
    }

    [Fact]
    public void Arrange_WhenHidden_MakesVisible()
    {
        var batch = _fixture.BatchToken;
        Assert.True(_commands.Hide(batch).Success);
        var before = _commands.GetInfo(batch);
        Assert.True(before.Success, before.ErrorMessage);
        Assert.False(before.IsVisible);

        var result = _commands.Arrange(batch, "center");

        Assert.True(result.Success);
        var after = _commands.GetInfo(batch);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.True(
            after.IsVisible,
            "Arrange should auto-show hidden Excel");
    }
}

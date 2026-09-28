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

        var result = _commands.Arrange(batch, preset);

        Assert.True(result.Success, $"Arrange '{preset}' failed: {result.ErrorMessage}");
        Assert.Equal("arrange", result.Action);
        Assert.Contains(preset, result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Arrange_InvalidPreset_Throws()
    {
        Assert.ThrowsAny<Exception>(() =>
            _commands.Arrange(_fixture.BatchToken, "invalid-preset"));
    }

    [Fact]
    public void Arrange_FullScreen_MaximizesWindow()
    {
        var batch = _fixture.BatchToken;

        var result = _commands.Arrange(batch, "full-screen");

        Assert.True(result.Success);
        var info = _commands.GetInfo(batch);
        Assert.True(info.IsVisible);
        Assert.Equal("maximized", info.WindowState);
    }

    [Fact]
    public void Arrange_WhenHidden_MakesVisible()
    {
        var batch = _fixture.BatchToken;
        _commands.Hide(batch);

        var result = _commands.Arrange(batch, "center");

        Assert.True(result.Success);
        Assert.True(
            _commands.GetInfo(batch).IsVisible,
            "Arrange should auto-show hidden Excel");
    }
}

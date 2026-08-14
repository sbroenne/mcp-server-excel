using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range hyperlinks operations
/// </summary>
public partial class RangeCommandsTests
{
    // === HYPERLINK OPERATIONS TESTS ===

    [Fact]
    public void AddHyperlink_CreatesHyperlink()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = _commands.AddHyperlink(
            batch,
            sheetName,
            "A1",
            "https://www.example.com",
            "Example Site",
            "Click to visit");

        // Assert
        Assert.True(result.Success);

        // Verify hyperlink exists
        var hyperlinkResult = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(hyperlinkResult.Success);
        Assert.Single(hyperlinkResult.Hyperlinks);
        // Excel normalizes URLs - may add trailing slash
        Assert.StartsWith("https://www.example.com", hyperlinkResult.Hyperlinks[0].Address);
    }

    [Fact]
    public void RemoveHyperlink_DeletesHyperlink()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.AddHyperlink(batch, sheetName, "A1", "https://www.example.com");

        // Act
        var result = _commands.RemoveHyperlink(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success);

        var hyperlinkResult = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.Empty(hyperlinkResult.Hyperlinks);
    }

    [Fact]
    public void ListHyperlinks_ReturnsAllHyperlinks()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.AddHyperlink(batch, sheetName, "A1", "https://site1.com");
        _commands.AddHyperlink(batch, sheetName, "B2", "https://site2.com");
        _commands.AddHyperlink(batch, sheetName, "C3", "https://site3.com");

        // Act
        var result = _commands.ListHyperlinks(batch, sheetName);

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.Hyperlinks.Count);
    }

    [Fact]
    public void AddHyperlink_InternalTarget_RoundTripsSubAddress()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var add = _commands.AddHyperlink(
            batch,
            sheetName,
            "A1",
            url: null,
            displayText: "Jump",
            tooltip: "Go to target",
            subAddress: $"'{sheetName}'!D5");
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        Assert.True(add.Success, add.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal($"'{sheetName}'!D5", hyperlink.SubAddress);
        Assert.Equal("Jump", hyperlink.DisplayText);
    }

    [Fact]
    public void UpdateHyperlink_ChangesTargetAndDisplayMetadata()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.AddHyperlink(batch, sheetName, "A1", "https://old.example.com", "Old");

        var update = _commands.UpdateHyperlink(
            batch,
            sheetName,
            "A1",
            url: "https://new.example.com",
            subAddress: "section",
            displayText: "New",
            tooltip: "Updated");
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        Assert.True(update.Success, update.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.StartsWith("https://new.example.com", hyperlink.Address);
        Assert.Equal("section", hyperlink.SubAddress);
        Assert.Equal("New", hyperlink.DisplayText);
        Assert.Equal("Updated", hyperlink.ScreenTip);
    }

    [Fact]
    public void ListHyperlinks_IncludesInternalTargetMetadata()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.AddHyperlink(batch, sheetName, "A1", url: null, subAddress: $"'{sheetName}'!B2");

        var list = _commands.ListHyperlinks(batch, sheetName);

        var hyperlink = Assert.Single(list.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal($"'{sheetName}'!B2", hyperlink.SubAddress);
    }

    [Fact]
    public void UpdateHyperlink_CannotRemoveOnlyTarget()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var originalTarget = $"'{sheetName}'!B2";
        _commands.AddHyperlink(batch, sheetName, "A1", url: null, subAddress: originalTarget);

        Assert.Throws<ArgumentException>(() =>
            _commands.UpdateHyperlink(batch, sheetName, "A1", subAddress: string.Empty));
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal(originalTarget, hyperlink.SubAddress);
    }
}


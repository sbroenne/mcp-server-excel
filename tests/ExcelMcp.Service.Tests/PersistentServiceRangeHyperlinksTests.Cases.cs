using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for range hyperlinks operations
/// </summary>
public sealed partial class PersistentServiceRangeHyperlinksTests
{
    // === HYPERLINK OPERATIONS TESTS ===

    [Fact]
    public void AddHyperlink_CreatesHyperlink()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = _commands.AddHyperlink(
            batch,
            sheetName,
            "A1",
            "https://www.example.com/",
            "Example Site",
            "Click to visit");

        // Assert
        Assert.True(result.Success);

        // Verify hyperlink exists
        var hyperlinkResult = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(hyperlinkResult.Success);
        var hyperlink = Assert.Single(hyperlinkResult.Hyperlinks);
        Assert.Equal("https://www.example.com/", hyperlink.Address);
        Assert.Equal("A1", hyperlink.CellAddress);
        Assert.Equal("Example Site", hyperlink.DisplayText);
        Assert.Equal("Click to visit", hyperlink.ScreenTip);
        Assert.False(hyperlink.IsInternal);
        var cells = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal("Example Site", Assert.Single(Assert.Single(cells.Values)));
    }

    [Fact]
    public void RemoveHyperlink_DeletesHyperlink()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1",
            "https://www.example.com/", "Retained text", "Original tooltip").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B1",
            "https://other.example.com/", "Untouched text", "Untouched tooltip").Success);

        // Act
        var result = _commands.RemoveHyperlink(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success);

        var hyperlinkResult = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(hyperlinkResult.Success, hyperlinkResult.ErrorMessage);
        Assert.Empty(hyperlinkResult.Hyperlinks);
        var cells = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["Retained text", "Untouched text"], Assert.Single(cells.Values));
        var remaining = _commands.ListHyperlinks(batch, sheetName);
        Assert.True(remaining.Success, remaining.ErrorMessage);
        var untouched = Assert.Single(remaining.Hyperlinks);
        Assert.Equal("B1", untouched.CellAddress);
        Assert.Equal("https://other.example.com/", untouched.Address);
        Assert.Equal("Untouched text", untouched.DisplayText);
        Assert.Equal("Untouched tooltip", untouched.ScreenTip);
    }

    [Fact]
    public void ListHyperlinks_ReturnsAllHyperlinks()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", "https://site1.com/", "First").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B2", "https://site2.com/", "Second").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "C3", "https://site3.com/", "Third").Success);

        // Act
        var result = _commands.ListHyperlinks(batch, sheetName);

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.Hyperlinks.Count);
        Assert.Equal(sheetName, result.SheetName);
        var hyperlinks = result.Hyperlinks.OrderBy(link => link.CellAddress, StringComparer.Ordinal).ToArray();
        Assert.Equal(["A1", "B2", "C3"], hyperlinks.Select(link => link.CellAddress));
        Assert.Equal(["https://site1.com/", "https://site2.com/", "https://site3.com/"],
            hyperlinks.Select(link => link.Address));
        Assert.Equal(["First", "Second", "Third"], hyperlinks.Select(link => link.DisplayText));
        Assert.All(hyperlinks, link => Assert.False(link.IsInternal));
    }

    [Fact]
    public void AddHyperlink_InternalTarget_RoundTripsSubAddress()
    {
        var batch = _fixture.BatchToken;
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
        Assert.True(get.Success, get.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal($"'{sheetName}'!D5", hyperlink.SubAddress);
        Assert.Equal("Jump", hyperlink.DisplayText);
        Assert.Equal("Go to target", hyperlink.ScreenTip);
        Assert.Equal("A1", hyperlink.CellAddress);
        Assert.Equal(string.Empty, hyperlink.Address);
    }

    [Fact]
    public void UpdateHyperlink_ChangesTargetAndDisplayMetadata()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", "https://old.example.com/", "Old").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B1", "https://other.example.com/", "Untouched").Success);
        var before = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(before.Success, before.ErrorMessage);
        Assert.Equal("https://old.example.com/", Assert.Single(before.Hyperlinks).Address);

        var update = _commands.UpdateHyperlink(
            batch,
            sheetName,
            "A1",
            url: "https://new.example.com/",
            subAddress: "section",
            displayText: "New",
            tooltip: "Updated");
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        Assert.True(update.Success, update.ErrorMessage);
        Assert.True(get.Success, get.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.Equal("https://new.example.com/", hyperlink.Address);
        Assert.Equal("section", hyperlink.SubAddress);
        Assert.Equal("New", hyperlink.DisplayText);
        Assert.Equal("Updated", hyperlink.ScreenTip);
        Assert.Equal("A1", hyperlink.CellAddress);
        var cells = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["New", "Untouched"], Assert.Single(cells.Values));
        var untouched = _commands.GetHyperlink(batch, sheetName, "B1");
        Assert.True(untouched.Success, untouched.ErrorMessage);
        Assert.Equal("https://other.example.com/", Assert.Single(untouched.Hyperlinks).Address);
    }

    [Fact]
    public void ListHyperlinks_IncludesInternalTargetMetadata()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", url: null,
            displayText: "Jump", tooltip: "Go to B2", subAddress: $"'{sheetName}'!B2").Success);

        var list = _commands.ListHyperlinks(batch, sheetName);

        Assert.True(list.Success, list.ErrorMessage);
        var hyperlink = Assert.Single(list.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal($"'{sheetName}'!B2", hyperlink.SubAddress);
        Assert.Equal("A1", hyperlink.CellAddress);
        Assert.Equal(string.Empty, hyperlink.Address);
        Assert.Equal("Jump", hyperlink.DisplayText);
        Assert.Equal("Go to B2", hyperlink.ScreenTip);
    }

    [Fact]
    public void UpdateHyperlink_CannotRemoveOnlyTarget()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var originalTarget = $"'{sheetName}'!B2";
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", url: null,
            displayText: "Original", tooltip: "Original tooltip", subAddress: originalTarget).Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B1", "https://other.example.com/", "Untouched").Success);
        var before = _commands.ListHyperlinks(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        var cellsBefore = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cellsBefore.Success, cellsBefore.ErrorMessage);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.UpdateHyperlink(batch, sheetName, "A1", subAddress: string.Empty,
                displayText: "Rejected text", tooltip: "Rejected tooltip"));
        Assert.Contains("must retain either an external address or an internal subAddress",
            exception.Message, StringComparison.Ordinal);
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        Assert.True(get.Success, get.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal(originalTarget, hyperlink.SubAddress);
        Assert.Equal("Original", hyperlink.DisplayText);
        Assert.Equal("Original tooltip", hyperlink.ScreenTip);
        var after = _commands.ListHyperlinks(batch, sheetName);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(before.Hyperlinks), JsonSerializer.Serialize(after.Hyperlinks));
        var cellsAfter = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cellsAfter.Success, cellsAfter.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(cellsBefore.Values), JsonSerializer.Serialize(cellsAfter.Values));
    }
}

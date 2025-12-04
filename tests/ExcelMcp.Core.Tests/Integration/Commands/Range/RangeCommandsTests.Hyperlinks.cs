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
}

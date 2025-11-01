using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Range;

/// <summary>
/// Tests for range hyperlinks operations
/// </summary>
public partial class RangeCommandsTests
{
    // === HYPERLINK OPERATIONS TESTS ===

    [Fact]
    public async Task AddHyperlinkAsync_CreatesHyperlink()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _commands.AddHyperlinkAsync(
            batch,
            "Sheet1",
            "A1",
            "https://www.example.com",
            "Example Site",
            "Click to visit");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);

        // Verify hyperlink exists
        var hyperlinkResult = await _commands.GetHyperlinkAsync(batch, "Sheet1", "A1");
        Assert.True(hyperlinkResult.Success);
        Assert.Single(hyperlinkResult.Hyperlinks);
        // Excel normalizes URLs - may add trailing slash
        Assert.StartsWith("https://www.example.com", hyperlinkResult.Hyperlinks[0].Address);
    }

    [Fact]
    public async Task RemoveHyperlinkAsync_DeletesHyperlink()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.AddHyperlinkAsync(batch, "Sheet1", "A1", "https://www.example.com");

        // Act
        var result = await _commands.RemoveHyperlinkAsync(batch, "Sheet1", "A1");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);

        var hyperlinkResult = await _commands.GetHyperlinkAsync(batch, "Sheet1", "A1");
        Assert.Empty(hyperlinkResult.Hyperlinks);
    }

    [Fact]
    public async Task ListHyperlinksAsync_ReturnsAllHyperlinks()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.AddHyperlinkAsync(batch, "Sheet1", "A1", "https://site1.com");
        await _commands.AddHyperlinkAsync(batch, "Sheet1", "B2", "https://site2.com");
        await _commands.AddHyperlinkAsync(batch, "Sheet1", "C3", "https://site3.com");

        // Act
        var result = await _commands.ListHyperlinksAsync(batch, "Sheet1");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.Hyperlinks.Count);
    }

}

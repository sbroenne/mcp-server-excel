using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range hyperlinks operations
/// </summary>
public partial class RangeCommandsTests
{
    /// <inheritdoc/>
    // === HYPERLINK OPERATIONS TESTS ===

    [Fact]
    public void AddHyperlink_CreatesHyperlink()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.AddHyperlink(
            batch,
            "Sheet1",
            "A1",
            "https://www.example.com",
            "Example Site",
            "Click to visit");
        // Assert
        Assert.True(result.Success);

        // Verify hyperlink exists
        var hyperlinkResult = _commands.GetHyperlink(batch, "Sheet1", "A1");
        Assert.True(hyperlinkResult.Success);
        Assert.Single(hyperlinkResult.Hyperlinks);
        // Excel normalizes URLs - may add trailing slash
        Assert.StartsWith("https://www.example.com", hyperlinkResult.Hyperlinks[0].Address);
    }
    /// <inheritdoc/>

    [Fact]
    public void RemoveHyperlink_DeletesHyperlink()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        _commands.AddHyperlink(batch, "Sheet1", "A1", "https://www.example.com");

        // Act
        var result = _commands.RemoveHyperlink(batch, "Sheet1", "A1");
        // Assert
        Assert.True(result.Success);

        var hyperlinkResult = _commands.GetHyperlink(batch, "Sheet1", "A1");
        Assert.Empty(hyperlinkResult.Hyperlinks);
    }
    /// <inheritdoc/>

    [Fact]
    public void ListHyperlinks_ReturnsAllHyperlinks()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), $"{Guid.NewGuid():N}", _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        _commands.AddHyperlink(batch, "Sheet1", "A1", "https://site1.com");
        _commands.AddHyperlink(batch, "Sheet1", "B2", "https://site2.com");
        _commands.AddHyperlink(batch, "Sheet1", "C3", "https://site3.com");

        // Act
        var result = _commands.ListHyperlinks(batch, "Sheet1");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.Hyperlinks.Count);
    }

}

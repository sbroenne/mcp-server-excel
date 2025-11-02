using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for PowerQuery name validation
/// Validates Excel's 80-character limit for Power Query names
/// </summary>
public partial class PowerQueryCommandsTests
{
    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_Import_EmptyQueryName_ReturnsError()
    {
        // Arrange
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Validation_Import_EmptyQueryName_ReturnsError));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ImportAsync(batch, "", testQueryFile, "connection-only");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_Import_WhitespaceQueryName_ReturnsError()
    {
        // Arrange
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Validation_Import_WhitespaceQueryName_ReturnsError));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ImportAsync(batch, "   ", testQueryFile, "connection-only");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Validation_Import_QueryNameExactly80Characters_ReturnsSuccess()
    {
        // Arrange - Create name with exactly 80 characters (Excel's actual limit)
        var queryName = new string('A', 80);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Validation_Import_QueryNameExactly80Characters_ReturnsSuccess));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile, "connection-only");

        // Assert
        Assert.True(result.Success, $"Expected success with 80-char name but got error: {result.ErrorMessage}");

        // Verify the query was actually created
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_Import_QueryName81Characters_ReturnsError()
    {
        // Arrange - Create name with 81 characters (exceeds Excel's limit)
        var queryName = new string('B', 81);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Validation_Import_QueryName81Characters_ReturnsError));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile, "connection-only");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("80-character limit", result.ErrorMessage);
        Assert.Contains("81", result.ErrorMessage);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_Update_QueryNameExceeds80Characters_ReturnsError()
    {
        // Arrange
        var longQueryName = new string('C', 100);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Validation_Update_QueryNameExceeds80Characters_ReturnsError));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.UpdateAsync(batch, longQueryName, testQueryFile);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("80-character limit", result.ErrorMessage);
        Assert.Contains("100", result.ErrorMessage);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_View_QueryNameExceeds80Characters_ReturnsError()
    {
        // Arrange
        var longQueryName = new string('D', 90);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ViewAsync(batch, longQueryName);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("80-character limit", result.ErrorMessage);
        Assert.Contains("90", result.ErrorMessage);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_Delete_QueryNameExceeds80Characters_ReturnsError()
    {
        // Arrange
        var longQueryName = new string('E', 95);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.DeleteAsync(batch, longQueryName);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("80-character limit", result.ErrorMessage);
        Assert.Contains("95", result.ErrorMessage);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_Refresh_QueryNameExceeds80Characters_ReturnsError()
    {
        // Arrange
        var longQueryName = new string('F', 85);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.RefreshAsync(batch, longQueryName);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("80-character limit", result.ErrorMessage);
        Assert.Contains("85", result.ErrorMessage);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_SetLoadToTableAsync_QueryNameExceeds80Characters_ReturnsError()
    {
        // Arrange
        var longQueryName = new string('G', 92);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.SetLoadToTableAsync(batch, longQueryName, "Sheet1");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("80-character limit", result.ErrorMessage);
        Assert.Contains("92", result.ErrorMessage);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public async Task Validation_GetLoadConfigAsync_QueryNameExceeds80Characters_ReturnsError()
    {
        // Arrange
        var longQueryName = new string('H', 88);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.GetLoadConfigAsync(batch, longQueryName);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("80-character limit", result.ErrorMessage);
        Assert.Contains("88", result.ErrorMessage);
    }
}

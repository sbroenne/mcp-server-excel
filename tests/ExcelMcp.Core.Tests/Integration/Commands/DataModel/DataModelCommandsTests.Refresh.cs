using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model refresh operations
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task Refresh_WithValidFile_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("Refresh_WithValidFile_ReturnsSuccessResult.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.RefreshAsync(batch);
        await batch.SaveAsync();

        // Assert - Demand success (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"Refresh MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");
    }

    [Fact]
    public async Task Refresh_WithRealisticDataModel_SucceedsOrIndicatesNoModel()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("Refresh_WithRealisticDataModel_SucceedsOrIndicatesNoModel.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.RefreshAsync(batch);
        await batch.SaveAsync();

        // Assert - Demand success (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"Refresh MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");

        // Verify refresh completed with correct file path
        Assert.NotNull(result.FilePath);
        Assert.Equal(testFile, result.FilePath);
    }
}

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
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.RefreshAsync(batch);
        await batch.SaveAsync();

        // Assert - Demand success (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"Refresh MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");
    }

    [Fact]
    public async Task Refresh_WithRealisticDataModel_SucceedsOrIndicatesNoModel()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.RefreshAsync(batch);
        await batch.SaveAsync();

        // Assert - Demand success (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"Refresh MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");
        
        // Verify refresh completed with correct file path
        Assert.NotNull(result.FilePath);
        Assert.Equal(_testExcelFile, result.FilePath);
    }
}

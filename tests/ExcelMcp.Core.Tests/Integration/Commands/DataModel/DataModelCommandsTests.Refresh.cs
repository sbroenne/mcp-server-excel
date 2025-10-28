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

        // Assert
        // Refresh should either succeed or indicate no Data Model
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");
    }

    [Fact]
    public async Task Refresh_WithRealisticDataModel_SucceedsOrIndicatesNoModel()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.RefreshAsync(batch);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        // If successful, should have refreshed the Data Model
        if (result.Success)
        {
            Assert.NotNull(result.FilePath);
            Assert.Equal(_testExcelFile, result.FilePath);
        }
    }
}

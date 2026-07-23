// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for Data Model Refresh operations.
/// Uses shared DataModelPivotTableFixture (non-destructive refresh).
/// </summary>
public partial class DataModelCommandsTests
{
    #region Refresh Tests

    /// <summary>
    /// Refreshes the entire Data Model.
    /// LLM use case: "refresh the data model"
    /// </summary>
    [Fact]
    public void Refresh_EntireModel_Succeeds()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Refresh(batch);

        Assert.True(result.Success, $"Refresh entire model failed: {result.ErrorMessage}");
        Assert.Equal(_dataModelFile, result.FilePath);
    }

    /// <summary>
    /// Refreshes a specific Data Model table by name.
    /// LLM use case: "refresh the SalesTable in the data model"
    /// </summary>
    [Fact]
    public void Refresh_SpecificTable_Succeeds()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Refresh(batch, tableName: "SalesTable");

        Assert.True(result.Success, $"Refresh specific table failed: {result.ErrorMessage}");
        Assert.Equal(_dataModelFile, result.FilePath);
    }

    /// <summary>
    /// Refreshing a non-existent table throws InvalidOperationException.
    /// LLM use case: error handling for typo in table name
    /// </summary>
    [Fact]
    public void Refresh_InvalidTableName_ThrowsInvalidOperationException()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var ex = Assert.Throws<InvalidOperationException>(
            () => _dataModelCommands.Refresh(batch, tableName: "NonExistentTable"));

        Assert.Contains("NonExistentTable", ex.Message);
    }

    /// <summary>
    /// Refreshing a workbook without a Data Model throws InvalidOperationException.
    /// LLM use case: error handling when data model doesn't exist
    /// </summary>
    [Fact]
    public void Refresh_NoDataModel_ThrowsInvalidOperationException()
    {
        // Create a fresh empty workbook (no Data Model)
        var tempDir = Path.Join(Path.GetTempPath(), $"RefreshTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        try
        {
            var emptyFile = CoreTestHelper.CreateUniqueTestFile(
                nameof(DataModelCommandsTests), nameof(Refresh_NoDataModel_ThrowsInvalidOperationException),
                tempDir);

            using var batch = ExcelSession.BeginBatch(emptyFile);

            Assert.Throws<InvalidOperationException>(
                () => _dataModelCommands.Refresh(batch));
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    /// <summary>
    /// Refresh with explicit timeout succeeds when within time limit.
    /// </summary>
    [Fact]
    public void Refresh_WithExplicitTimeout_Succeeds()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = _dataModelCommands.Refresh(batch, timeout: TimeSpan.FromMinutes(5));

        Assert.True(result.Success, $"Refresh with timeout failed: {result.ErrorMessage}");
    }

    #endregion
}

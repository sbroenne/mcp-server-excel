// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Data Model Refresh operations.
/// Uses shared DataModelPivotTableFixture (non-destructive refresh).
/// </summary>
public partial class PersistentServiceDataModelCommandsTests
{
    #region Refresh Tests

    /// <summary>
    /// Refreshes the entire Data Model.
    /// LLM use case: "refresh the data model"
    /// </summary>
    [Fact]
    public void Refresh_EntireModel_Succeeds()
    {
        var batch = _fixture.BatchToken;
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
        var batch = _fixture.BatchToken;
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
        var batch = _fixture.BatchToken;

        var ex = Assert.Throws<InvalidOperationException>(
            () => _dataModelCommands.Refresh(batch, tableName: "NonExistentTable"));

        Assert.Contains("NonExistentTable", ex.Message);
    }

    /// <summary>
    /// Refresh with explicit timeout succeeds when within time limit.
    /// </summary>
    [Fact]
    public void Refresh_WithExplicitTimeout_Succeeds()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Refresh(batch, timeout: TimeSpan.FromMinutes(5));

        Assert.True(result.Success, $"Refresh with timeout failed: {result.ErrorMessage}");
    }

    #endregion
}

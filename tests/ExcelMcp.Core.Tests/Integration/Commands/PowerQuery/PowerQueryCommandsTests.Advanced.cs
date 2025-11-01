using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Advanced PowerQuery tests: error handling and query references
/// </summary>
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Verifies error handling when operating on non-existent queries.
    /// Tests that all operations return appropriate "not found" errors.
    /// </summary>
    [Fact]
    public async Task Operations_WithNonExistentQuery_ReturnNotFoundError()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(PowerQueryCommandsTests), nameof(Operations_WithNonExistentQuery_ReturnNotFoundError), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        // Act & Assert - Test multiple operations return "not found" error
        var viewResult = await _powerQueryCommands.ViewAsync(batch, "NonExistentQuery");
        Assert.False(viewResult.Success);
        Assert.Contains("not found", viewResult.ErrorMessage);

        var getConfigResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "NonExistentQuery");
        Assert.False(getConfigResult.Success);
        Assert.Contains("not found", getConfigResult.ErrorMessage);

        var setTableResult = await _powerQueryCommands.SetLoadToTableAsync(batch, "NonExistentQuery", "TestSheet");
        Assert.False(setTableResult.Success);
        Assert.Contains("not found", setTableResult.ErrorMessage);

        var setModelResult = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "NonExistentQuery");
        Assert.False(setModelResult.Success);
        Assert.Contains("not found", setModelResult.ErrorMessage);

        var setBothResult = await _powerQueryCommands.SetLoadToBothAsync(batch, "NonExistentQuery", "TestSheet");
        Assert.False(setBothResult.Success);
        Assert.Contains("not found", setBothResult.ErrorMessage);

        var setConnResult = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "NonExistentQuery");
        Assert.False(setConnResult.Success);
        Assert.Contains("not found", setConnResult.ErrorMessage);
    }

    /// <summary>
    /// Verifies that a Power Query can successfully reference and load data from another Power Query.
    /// Tests the common pattern where one query (DerivedQuery) references another query (SourceQuery).
    /// </summary>
    [Fact]
    public async Task Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(PowerQueryCommandsTests), nameof(Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully), _tempDir);

        // Create M code for the source query (base data)
        string sourceQueryMCode = @"let
    Source = #table(
        {""ProductID"", ""ProductName"", ""Price""},
        {
            {1, ""Widget"", 10.99},
            {2, ""Gadget"", 25.50},
            {3, ""Doohickey"", 15.75}
        }
    )
in
    Source";

        var sourceQueryFile = Path.Combine(_tempDir, $"SourceQuery_{Guid.NewGuid():N}.pq");
        File.WriteAllText(sourceQueryFile, sourceQueryMCode);

        // Create M code for the derived query (references the source query)
        string derivedQueryMCode = @"let
    Source = SourceQuery,
    FilteredRows = Table.SelectRows(Source, each [Price] > 15)
in
    FilteredRows";

        var derivedQueryFile = Path.Combine(_tempDir, $"DerivedQuery_{Guid.NewGuid():N}.pq");
        File.WriteAllText(derivedQueryFile, derivedQueryMCode);

        // Act & Assert
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        // Import source query first
        var sourceImportResult = await _powerQueryCommands.ImportAsync(
            batch,
            "SourceQuery",
            sourceQueryFile,
            loadDestination: "worksheet");

        Assert.True(sourceImportResult.Success,
            $"Source query import failed: {sourceImportResult.ErrorMessage}");

        // Import derived query (references SourceQuery)
        var derivedImportResult = await _powerQueryCommands.ImportAsync(
            batch,
            "DerivedQuery",
            derivedQueryFile,
            loadDestination: "worksheet");

        Assert.True(derivedImportResult.Success,
            $"Derived query import failed: {derivedImportResult.ErrorMessage}");

        // Verify both queries exist in the workbook
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Equal(2, listResult.Queries.Count);
        Assert.Contains(listResult.Queries, q => q.Name == "SourceQuery");
        Assert.Contains(listResult.Queries, q => q.Name == "DerivedQuery");

        // Verify the source query M code
        var sourceViewResult = await _powerQueryCommands.ViewAsync(batch, "SourceQuery");
        Assert.True(sourceViewResult.Success);
        Assert.Equal("SourceQuery", sourceViewResult.QueryName);
        Assert.Contains("#table", sourceViewResult.MCode);
        Assert.Contains("ProductID", sourceViewResult.MCode);

        // Verify the derived query M code references SourceQuery
        var derivedViewResult = await _powerQueryCommands.ViewAsync(batch, "DerivedQuery");
        Assert.True(derivedViewResult.Success);
        Assert.Equal("DerivedQuery", derivedViewResult.QueryName);
        Assert.Contains("SourceQuery", derivedViewResult.MCode);
        Assert.Contains("Table.SelectRows", derivedViewResult.MCode);
        Assert.Contains("Price", derivedViewResult.MCode);

        // Refresh both queries to ensure they execute successfully
        var sourceRefreshResult = await _powerQueryCommands.RefreshAsync(batch, "SourceQuery");
        Assert.True(sourceRefreshResult.Success,
            $"Source query refresh failed: {sourceRefreshResult.ErrorMessage}");

        var derivedRefreshResult = await _powerQueryCommands.RefreshAsync(batch, "DerivedQuery");
        Assert.True(derivedRefreshResult.Success,
            $"Derived query refresh failed: {derivedRefreshResult.ErrorMessage}");

        await batch.SaveAsync();
    }
}

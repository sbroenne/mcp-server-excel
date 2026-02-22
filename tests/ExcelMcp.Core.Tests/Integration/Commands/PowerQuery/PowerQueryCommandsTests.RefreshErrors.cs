using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for Power Query Refresh error propagation.
///
/// BUG FIXED (Issue #399): Refresh was returning Success=true even when Power Query had formula errors.
///
/// ROOT CAUSE: The original code used Connection.Refresh() which silently swallows errors
/// for worksheet queries (InModel=false). Only QueryTable.Refresh(false) throws actual
/// Power Query formula errors for worksheet queries.
///
/// COM API BEHAVIOR:
/// | Connection Type | InModel | Error Thrown By |
/// |-----------------|---------|-----------------|
/// | Worksheet       | false   | QueryTable.Refresh(false) |
/// | Data Model      | true    | Connection.Refresh() |
///
/// These tests verify that errors are now properly propagated for both connection types.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Regression test for Issue #399: Worksheet query with invalid M code should throw during refresh.
    /// This tests the QueryTable.Refresh() path for worksheet queries (InModel=false).
    ///
    /// Setup: Create a valid query first (which loads successfully), then update it to have
    /// invalid M code, then verify that Refresh throws the error.
    /// </summary>
    [Fact]
    public void Refresh_WorksheetQueryWithInvalidMCode_ThrowsError()
    {
        // Arrange - Create a worksheet query with VALID M code first
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "BrokenWorksheetQuery";

        // Valid M code for initial creation
        var validMCode = @"let
    Source = #table({""X""}, {{1}})
in
    Source";

        // Invalid M code that will fail on refresh
        var invalidMCode = @"let
    Source = NonExistentFunction()
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query loaded to worksheet with VALID code (should succeed)
        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.LoadToTable);
        batch.Save();

        // Now UPDATE the query to have invalid M code (WITHOUT auto-refresh)
        _powerQueryCommands.Update(batch, queryName, invalidMCode, refresh: false);
        batch.Save();

        // Act & Assert - Refresh should throw with the Power Query error message
        var exception = Assert.ThrowsAny<Exception>(() =>
            _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(1)));

        // Verify error message contains Power Query error details
        Assert.Contains("Expression.Error", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Regression test for Issue #399: Query referencing non-existent table should throw during refresh.
    /// This is the exact scenario from the original bug report (Milestone_Export query).
    ///
    /// Setup: Create a valid query first, then update it to reference a non-existent table.
    /// </summary>
    [Fact]
    public void Refresh_QueryReferencingNonExistentTable_ThrowsError()
    {
        // Arrange - Create a query that starts valid, then we'll change it
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "TableRefQuery";

        // Valid M code for initial creation
        var validMCode = @"let
    Source = #table({""X""}, {{1}})
in
    Source";

        // Invalid M code that references a table that doesn't exist in the workbook
        var invalidMCode = @"let
    Source = Excel.CurrentWorkbook(){[Name=""NonExistentTable""]}[Content]
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query loaded to worksheet with VALID code
        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.LoadToTable);
        batch.Save();

        // Update to reference non-existent table (WITHOUT auto-refresh)
        _powerQueryCommands.Update(batch, queryName, invalidMCode, refresh: false);
        batch.Save();

        // Act & Assert - Refresh should throw
        var exception = Assert.ThrowsAny<Exception>(() =>
            _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(1)));

        // Error message should indicate the table wasn't found
        Assert.True(
            exception.Message.Contains("Expression.Error", StringComparison.OrdinalIgnoreCase) ||
            exception.Message.Contains("DataSource.Error", StringComparison.OrdinalIgnoreCase) ||
            exception.Message.Contains("didn't find", StringComparison.OrdinalIgnoreCase),
            $"Expected Power Query error but got: {exception.Message}");
    }

    /// <summary>
    /// Regression test: Valid worksheet query should refresh successfully (no false positives).
    /// </summary>
    [Fact]
    public void Refresh_ValidWorksheetQuery_Succeeds()
    {
        // Arrange - Create a valid worksheet query
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "ValidWorksheetQuery";

        var validMCode = @"let
    Source = #table(
        {""Name"", ""Value""},
        {{""Test"", 100}, {""Data"", 200}}
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query loaded to worksheet
        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.LoadToTable);
        batch.Save();

        // Act - Refresh should succeed
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(1));

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
        // Note: LoadedToSheet is populated during Create, not necessarily during Refresh
    }

    /// <summary>
    /// Regression test for Issue #399: Data Model query with invalid M code should throw during refresh.
    /// This tests the Connection.Refresh() path for Data Model queries (InModel=true).
    ///
    /// Setup: Create a valid Data Model query first, then update it to have invalid M code.
    /// </summary>
    [Fact]
    public void Refresh_DataModelQueryWithInvalidMCode_ThrowsError()
    {
        // Arrange - Create a Data Model query with VALID M code first
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "BrokenDataModelQuery";

        // Valid M code for initial creation
        var validMCode = @"let
    Source = #table({""X""}, {{1}})
in
    Source";

        // Invalid M code that will fail during refresh
        var invalidMCode = @"let
    Source = UndefinedReference
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query loaded to Data Model with VALID code (should succeed)
        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.LoadToDataModel);
        batch.Save();

        // Now UPDATE the query to have invalid M code (WITHOUT auto-refresh)
        _powerQueryCommands.Update(batch, queryName, invalidMCode, refresh: false);
        batch.Save();

        // Act & Assert - Refresh should throw with Power Query error
        var exception = Assert.ThrowsAny<Exception>(() =>
            _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(1)));

        // Verify error message contains Power Query error details or Data Model error
        Assert.True(
            exception.Message.Contains("Expression.Error", StringComparison.OrdinalIgnoreCase) ||
            exception.Message.Contains("Data Model", StringComparison.OrdinalIgnoreCase) ||
            exception.Message.Contains("couldn't get data", StringComparison.OrdinalIgnoreCase),
            $"Expected Power Query/Data Model error but got: {exception.Message}");
    }

    /// <summary>
    /// Regression test: Valid Data Model query should refresh successfully.
    /// </summary>
    [Fact]
    public void Refresh_ValidDataModelQuery_Succeeds()
    {
        // Arrange - Create a valid Data Model query
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "ValidDataModelQuery";

        var validMCode = @"let
    Source = #table(
        {""Category"", ""Amount""},
        {{""Sales"", 1000}, {""Marketing"", 500}}
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create query loaded to Data Model
        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.LoadToDataModel);
        batch.Save();

        // Act - Refresh should succeed
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(1));

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
    }

    /// <summary>
    /// Regression test: Refresh with TimeSpan.Zero timeout should use the default 5-minute timeout
    /// instead of throwing ArgumentOutOfRangeException.
    ///
    /// BUG: The source generator produces `args.Timeout ?? default(TimeSpan)` = TimeSpan.Zero
    /// when no timeout is supplied (CLI without --timeout, MCP without timeout parameter).
    /// The Core method used to throw ArgumentOutOfRangeException("Timeout must be greater than zero.").
    /// FIX: TimeSpan.Zero now falls back to 5-minute default.
    /// </summary>
    [Fact]
    public void Refresh_ZeroTimeout_UsesDefaultAndSucceeds()
    {
        // Arrange - Create a valid worksheet query
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "ZeroTimeoutQuery";

        var validMCode = @"let
    Source = #table({""Name""}, {{""Test""}})
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.LoadToTable);
        batch.Save();

        // Act - Refresh with TimeSpan.Zero (the exact value generated when timeout is omitted)
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.Zero);

        // Assert - Should succeed, not throw ArgumentOutOfRangeException
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
    }

    /// <summary>
    /// Regression test: Refresh with negative timeout should use the default 5-minute timeout.
    /// </summary>
    [Fact]
    public void Refresh_NegativeTimeout_UsesDefaultAndSucceeds()
    {
        // Arrange - Create a valid worksheet query
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "NegativeTimeoutQuery";

        var validMCode = @"let
    Source = #table({""Name""}, {{""Test""}})
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.LoadToTable);
        batch.Save();

        // Act - Refresh with negative timeout
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromSeconds(-1));

        // Assert - Should succeed, not throw ArgumentOutOfRangeException
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
    }

    /// <summary>
    /// Verifies that refresh throws when query is connection-only (no QueryTable or DataModel connection).
    /// Connection-only queries have no mechanism to refresh - they're only query definitions.
    /// </summary>
    [Fact]
    public void Refresh_ConnectionOnlyQuery_ThrowsBecauseNoRefreshMechanism()
    {
        // Arrange - Create a connection-only query (no worksheet table, no Data Model)
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "ConnectionOnlyQuery";

        var validMCode = @"let
    Source = #table({""X""}, {{1}})
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create connection-only query
        _powerQueryCommands.Create(batch, queryName, validMCode, PowerQueryLoadMode.ConnectionOnly);
        batch.Save();

        // Act & Assert - Connection-only queries cannot be refreshed because there's
        // no QueryTable (worksheet) or InModel=true connection (Data Model) to refresh
        var exception = Assert.ThrowsAny<Exception>(() =>
            _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(1)));

        // Should indicate no refresh mechanism found
        Assert.Contains("Could not find connection or table", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}





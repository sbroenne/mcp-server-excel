using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Regression tests for the BackgroundQuery CPU spin bug in Power Query refresh.
///
/// Root cause: PowerQueryCommands.RefreshWorkbookConnection (in Helpers.cs) forced
/// the OLEDBConnection.BackgroundQuery to true before calling connection.Refresh().
/// With BackgroundQuery = true, connection.Refresh() returned immediately (async),
/// leaving the STA thread to poll connection.Refreshing with Thread.Sleep(200).
/// OleMessageFilter caused Thread.Sleep to return immediately on every COM event
/// (via MsgWaitForMultipleObjectsEx), creating a 100% CPU spin for the full
/// duration of the refresh — seconds to minutes per query.
///
/// Fix: Force BackgroundQuery = false before calling connection.Refresh().
/// Refresh() then blocks synchronously; connection.Refreshing is false on return;
/// the polling loop exits in 0 iterations. Zero spin.
///
/// This class tests the Connection.Refresh() path (Strategy 2 in RefreshConnectionByQueryName),
/// which is used for Data Model queries. Worksheet queries use QueryTable.Refresh() (Strategy 1).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public class PowerQueryBackgroundQueryRegressionTests : IClassFixture<TempDirectoryFixture>
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly TempDirectoryFixture _fixture;

    public PowerQueryBackgroundQueryRegressionTests(TempDirectoryFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _fixture = fixture;
    }

    /// <summary>
    /// Regression test: A Power Query loaded to the Data Model must refresh successfully
    /// via the Connection.Refresh() path without causing a CPU spin.
    ///
    /// This tests Strategy 2 of RefreshConnectionByQueryName — the path where
    /// RefreshWorkbookConnection() was calling BackgroundQuery = true and spinning.
    /// </summary>
    [Fact]
    public void Refresh_DataModelQuery_CompletesWithoutCpuSpin()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "DM_BGQ_Regression_" + Guid.NewGuid().ToString("N")[..8];

        const string mCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Value""},
        {
            {1, ""Alpha"", 100},
            {2, ""Beta"", 200},
            {3, ""Gamma"", 300}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create the query and load to Data Model (uses Connection.Refresh() path internally)
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToDataModel);

        // Act — explicitly refresh via PowerQueryCommands.Refresh() which calls
        // RefreshConnectionByQueryName → Strategy 2 → RefreshWorkbookConnection.
        // Before the fix: this would spin at ~100% CPU for the duration of the refresh.
        // After the fix: BackgroundQuery is forced to false, connection.Refresh() blocks
        // synchronously, the polling loop exits immediately (0 iterations).
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(2));

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.Equal(queryName, result.QueryName);
        // Not loaded to worksheet (Data Model only)
        Assert.True(
            result.IsConnectionOnly || string.IsNullOrEmpty(result.LoadedToSheet),
            "Data Model query should not be loaded to a worksheet.");
    }

    /// <summary>
    /// Regression test: A worksheet Power Query must also refresh without a CPU spin.
    /// This tests Strategy 1 of RefreshConnectionByQueryName (QueryTable.Refresh path).
    /// </summary>
    [Fact]
    public void Refresh_WorksheetQuery_CompletesSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "WS_BGQ_Regression_" + Guid.NewGuid().ToString("N")[..8];

        const string mCode = @"let
    Source = #table(
        {""ID"", ""Name""},
        {{1, ""Alpha""}, {2, ""Beta""}}
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create the query loaded to worksheet (uses QueryTable.Refresh() path)
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable);

        // Act — refresh via Strategy 1 path (QueryTable.Refresh)
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(2));

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(string.IsNullOrEmpty(result.LoadedToSheet),
            "Worksheet query should have a loaded sheet.");
    }
}

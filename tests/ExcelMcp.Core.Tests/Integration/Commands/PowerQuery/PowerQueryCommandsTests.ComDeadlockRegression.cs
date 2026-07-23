using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Regression tests for the COM deadlock bug introduced by <c>EnterLongOperation</c>
/// in the PowerQuery refresh path (fixed in v1.8.29).
///
/// ROOT CAUSE:
/// <c>EnterLongOperation()</c> set <c>_isInLongOperation=true</c>, which caused
/// <c>IMessageFilter.HandleInComingCall</c> to return <c>SERVERCALL_RETRYLATER</c>
/// for ALL inbound COM calls — including the essential MashupHost callbacks that
/// Excel needs to deliver query results back to our STA thread during a synchronous
/// <c>QueryTable.Refresh(false)</c> or <c>connection.Refresh()</c>.
///
/// FAILURE MODE (if deadlock reintroduced):
/// The test will hang on <c>Refresh()</c> for the full timeout and then throw.
/// The test MUST fail loudly rather than passing slowly.
///
/// WHAT THIS GUARDS:
/// - <c>queryTable.Refresh(false)</c> path (worksheet-loaded query, <c>InModel=false</c>)
/// - <c>connection.Refresh()</c> path (data model query, <c>InModel=true</c>)
///
/// KNOWN TRADE-OFF:
/// CPU during refresh is ~88% (vs ~25% when <c>EnterLongOperation</c> was active).
/// Elevated CPU is accepted as preferable to a permanent hang.
/// See <c>PowerQueryRefreshCpuSpinTests</c> for the intentionally-failing CPU tests.
/// </summary>
public partial class PowerQueryCommandsTests
{
    private static readonly string DeadlockRegressionMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Value""},
        {
            {1, ""Alpha"",  100},
            {2, ""Beta"",   200},
            {3, ""Gamma"",  300}
        }
    )
in
    Source";

    /// <summary>
    /// Regression test: worksheet-loaded query (<c>LoadToTable</c>, <c>InModel=false</c>)
    /// must complete refresh within a strict timeout.
    ///
    /// This guards the <c>queryTable.Refresh(false)</c> code path in
    /// <c>PowerQueryCommands.Helpers.cs / RefreshQueryTableByName</c>.
    ///
    /// If <c>EnterLongOperation</c> is re-added before <c>queryTable.Refresh(false)</c>,
    /// this test will hang for the full 90-second timeout and then fail with
    /// <c>OperationCanceledException</c> — making the regression immediately visible.
    /// </summary>
    [Fact]
    public void Refresh_WorksheetLoadedQuery_CompletesWithoutDeadlock()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "DR_WorksheetQuery"; // ≤31 chars (Excel sheet name limit)

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create as worksheet-loaded (LoadToTable → InModel=false → queryTable.Refresh path)
        // This is the exact load destination that triggered the 30-minute hang in production.
        _powerQueryCommands.Create(batch, queryName, DeadlockRegressionMCode, PowerQueryLoadMode.LoadToTable);
        batch.Save();

        var sw = Stopwatch.StartNew();

        // Act — strict 90-second timeout.
        // A healthy refresh of this trivial in-memory table completes in ~5-15s.
        // If deadlock is reintroduced this will hang until timeout and throw.
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromSeconds(90));

        sw.Stop();

        // Assert
        Assert.True(result.Success, $"Worksheet refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors, $"Refresh reported errors: {string.Join(", ", result.ErrorMessages)}");

        // Sanity check: completing in under 60s means we were NOT deadlocked.
        // (Under deadlock the call blocks until the timeout fires.)
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
            $"Refresh took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
            $"Possible deadlock regression: check if EnterLongOperation was re-added " +
            $"before queryTable.Refresh(false) in PowerQueryCommands.Helpers.cs.");
    }

    /// <summary>
    /// Regression test: data model query (<c>LoadToDataModel</c>, <c>InModel=true</c>)
    /// must complete refresh within a strict timeout.
    ///
    /// This guards the <c>connection.Refresh()</c> code path in
    /// <c>PowerQueryCommands.Helpers.cs / RefreshConnectionByQueryName</c>.
    ///
    /// If <c>EnterLongOperation</c> is re-added before <c>connection.Refresh()</c>,
    /// this test will hang for the full 90-second timeout and then fail.
    /// </summary>
    [Fact]
    public void Refresh_DataModelLoadedQuery_CompletesWithoutDeadlock()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "DR_DataModelQuery"; // ≤31 chars

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create as data model query (LoadToDataModel → InModel=true → connection.Refresh path)
        _powerQueryCommands.Create(batch, queryName, DeadlockRegressionMCode, PowerQueryLoadMode.LoadToDataModel);
        batch.Save();

        var sw = Stopwatch.StartNew();

        // Act — strict 90-second timeout.
        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromSeconds(90));

        sw.Stop();

        // Assert
        Assert.True(result.Success, $"Data model refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors, $"Refresh reported errors: {string.Join(", ", result.ErrorMessages)}");

        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
            $"Refresh took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
            $"Possible deadlock regression: check if EnterLongOperation was re-added " +
            $"before connection.Refresh() in PowerQueryCommands.Helpers.cs.");
    }

    /// <summary>
    /// Regression test: Evaluate uses a temporary worksheet QueryTable refresh internally.
    /// That path must not reintroduce the same COM deadlock fixed in Refresh().
    /// </summary>
    [Fact]
    public void Evaluate_TemporaryWorksheetQuery_CompletesWithoutDeadlock()
    {
        var testExcelFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(90),
            testExcelFile);

        var sw = Stopwatch.StartNew();

        var result = _powerQueryCommands.Evaluate(batch, DeadlockRegressionMCode);

        sw.Stop();

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(3, result.RowCount);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
            $"Evaluate took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
            $"Possible deadlock regression in the temporary QueryTable.Refresh(false) path.");
    }

    /// <summary>
    /// Regression test: LoadToTable applies destination changes by creating a worksheet QueryTable
    /// and refreshing it synchronously. That path must not deadlock.
    /// </summary>
    [Fact]
    public void LoadTo_Table_CompletesWithoutDeadlock()
    {
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "DR_LoadToTable";

        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(90),
            testExcelFile);

        _powerQueryCommands.Create(batch, queryName, DeadlockRegressionMCode, PowerQueryLoadMode.ConnectionOnly);
        batch.Save();

        var sw = Stopwatch.StartNew();

        var result = _powerQueryCommands.LoadTo(batch, queryName, PowerQueryLoadMode.LoadToTable, queryName, "A1");

        sw.Stop();

        Assert.True(result.Success, $"LoadTo(Table) failed: {result.ErrorMessage}");
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
            $"LoadTo(Table) took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
            $"Possible deadlock regression in the worksheet QueryTable refresh path.");
    }

    /// <summary>
    /// Regression test: LoadToDataModel creates a workbook connection and refreshes it synchronously.
    /// That connection.Refresh() path must not deadlock.
    /// </summary>
    [Fact]
    public void LoadTo_DataModel_CompletesWithoutDeadlock()
    {
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "DR_LoadToModel";

        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(90),
            testExcelFile);

        _powerQueryCommands.Create(batch, queryName, DeadlockRegressionMCode, PowerQueryLoadMode.ConnectionOnly);
        batch.Save();

        var sw = Stopwatch.StartNew();

        var result = _powerQueryCommands.LoadTo(batch, queryName, PowerQueryLoadMode.LoadToDataModel);

        sw.Stop();

        Assert.True(result.Success, $"LoadTo(DataModel) failed: {result.ErrorMessage}");
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
            $"LoadTo(DataModel) took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
            $"Possible deadlock regression in the data model connection.Refresh() path.");
    }

    /// <summary>
    /// Regression test: Update on a worksheet-loaded query must reuse the same COM-safe
    /// synchronous refresh semantics as Refresh().
    /// </summary>
    [Fact]
    public void Update_WorksheetLoadedQuery_CompletesWithoutDeadlock()
    {
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "DR_UpdateSheet";
        var updatedMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Value"", ""Extra""},
        {
            {1, ""Alpha"", 100, ""A""},
            {2, ""Beta"", 200, ""B""},
            {3, ""Gamma"", 300, ""C""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(90),
            testExcelFile);

        _powerQueryCommands.Create(batch, queryName, DeadlockRegressionMCode, PowerQueryLoadMode.LoadToTable);
        batch.Save();

        var sw = Stopwatch.StartNew();

        var result = _powerQueryCommands.Update(batch, queryName, updatedMCode, refresh: true);

        sw.Stop();

        Assert.True(result.Success, $"Update(worksheet) failed: {result.ErrorMessage}");
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
            $"Update(worksheet) took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
            $"Possible deadlock regression in the QueryTable.Refresh(false) update path.");
    }

    /// <summary>
    /// Regression test: Update on a data-model query must reuse the same COM-safe
    /// synchronous refresh semantics as Refresh().
    /// </summary>
    [Fact]
    public void Update_DataModelLoadedQuery_CompletesWithoutDeadlock()
    {
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "DR_UpdateModel";
        var updatedMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Value"", ""Extra""},
        {
            {1, ""Alpha"", 100, ""A""},
            {2, ""Beta"", 200, ""B""},
            {3, ""Gamma"", 300, ""C""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(
            show: false,
            operationTimeout: TimeSpan.FromSeconds(90),
            testExcelFile);

        _powerQueryCommands.Create(batch, queryName, DeadlockRegressionMCode, PowerQueryLoadMode.LoadToDataModel);
        batch.Save();

        var sw = Stopwatch.StartNew();

        var result = _powerQueryCommands.Update(batch, queryName, updatedMCode, refresh: true);

        sw.Stop();

        Assert.True(result.Success, $"Update(data model) failed: {result.ErrorMessage}");
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
            $"Update(data model) took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
            $"Possible deadlock regression in the connection.Refresh() update path.");
    }
}

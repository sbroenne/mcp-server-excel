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
}

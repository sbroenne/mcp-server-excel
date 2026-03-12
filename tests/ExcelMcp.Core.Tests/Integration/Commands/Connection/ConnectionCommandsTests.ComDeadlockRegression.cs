using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Regression tests for the COM deadlock class caused by rejecting inbound callbacks
/// during synchronous connection load/refresh operations.
///
/// ROOT CAUSE:
/// Some connection paths wrapped <c>QueryTable.Refresh(false)</c> or <c>connection.Refresh()</c>
/// in <c>EnterLongOperation()</c>. That caused <c>HandleInComingCall</c> to reject inbound COM
/// callbacks with <c>SERVERCALL_RETRYLATER</c>. Excel/providers may need those callbacks to complete
/// a synchronous refresh, so the call can deadlock until the timeout fires.
///
/// WHAT THIS GUARDS:
/// - <c>ConnectionCommands.LoadTo()</c> worksheet QueryTable refresh path
/// - <c>ConnectionCommands.Refresh()</c> workbook connection refresh path
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests
{
    /// <summary>
    /// Regression test: worksheet load of a connection must complete within a strict timeout.
    /// If inbound callback rejection is reintroduced, this will hang until the timeout fires.
    /// </summary>
    [Fact]
    public void LoadTo_AceOleDbConnection_CompletesWithoutDeadlock()
    {
        var (testFile, sourceWorkbook, connectionName) = SetupAceOleDbConnection();

        try
        {
            using var batch = ExcelSession.BeginBatch(
                show: false,
                operationTimeout: TimeSpan.FromSeconds(90),
                testFile);

            var sw = Stopwatch.StartNew();

            _commands.LoadTo(batch, connectionName, "ProductsData");

            sw.Stop();

            Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
                $"LoadTo took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
                "Possible deadlock regression in the connection QueryTable.Refresh(false) path.");
        }
        finally
        {
            if (System.IO.File.Exists(sourceWorkbook))
            {
                System.IO.File.Delete(sourceWorkbook);
            }
        }
    }

    /// <summary>
    /// Regression test: synchronous connection refresh must complete within a strict timeout.
    /// The load step runs first to materialize the destination table; the refresh itself is the path under test.
    /// </summary>
    [Fact]
    public void Refresh_AceOleDbConnection_CompletesWithoutDeadlock()
    {
        var (testFile, sourceWorkbook, connectionName) = SetupAceOleDbConnection();

        try
        {
            using (var loadBatch = ExcelSession.BeginBatch(
                show: false,
                operationTimeout: TimeSpan.FromSeconds(90),
                testFile))
            {
                _commands.LoadTo(loadBatch, connectionName, "ProductsData");
                loadBatch.Save();
            }

            using var refreshBatch = ExcelSession.BeginBatch(
                show: false,
                operationTimeout: TimeSpan.FromSeconds(90),
                testFile);

            var sw = Stopwatch.StartNew();

            _commands.Refresh(refreshBatch, connectionName);

            sw.Stop();

            Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60),
                $"Refresh took {sw.Elapsed.TotalSeconds:F1}s — suspiciously slow. " +
                "Possible deadlock regression in the connection.Refresh() path.");
        }
        finally
        {
            if (System.IO.File.Exists(sourceWorkbook))
            {
                System.IO.File.Delete(sourceWorkbook);
            }
        }
    }
}
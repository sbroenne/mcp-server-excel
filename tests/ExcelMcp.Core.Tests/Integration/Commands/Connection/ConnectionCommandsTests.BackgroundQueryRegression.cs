using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Regression tests for the BackgroundQuery CPU spin fix.
///
/// Root cause: RefreshWorkbookConnection originally forced BackgroundQuery = true before
/// calling connection.Refresh(). With BackgroundQuery = true, connection.Refresh() returns
/// immediately (async mode). The STA thread then polls connection.Refreshing with
/// Thread.Sleep(200). Because OleMessageFilter is registered on the STA thread,
/// COM events from Excel during the background refresh cause MsgWaitForMultipleObjectsEx
/// to wake Thread.Sleep immediately — turning the polling loop into a 100% CPU spin
/// lasting the full duration of the refresh.
///
/// Fix: Force BackgroundQuery = false before calling connection.Refresh() so it blocks
/// the STA thread synchronously until done. The polling loop then exits immediately with
/// zero iterations (connection.Refreshing is already false on return).
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests
{
    /// <summary>
    /// Regression test: BackgroundQuery must be preserved (restored) after refresh.
    /// When BackgroundQuery is true, the fix temporarily sets it to false (sync refresh),
    /// then restores it. This test verifies the restore actually happens.
    /// </summary>
    [Fact]
    public void Refresh_BackgroundQueryTrue_RestoredAfterRefresh()
    {
        var (testFile, sourceWorkbook, connectionName) = SetupAceOleDbConnection();

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);
            _commands.LoadTo(batch, connectionName, "ProductsData");

            // Verify BackgroundQuery starts as true (set by ConnectionTestHelper)
            var preBefore = _commands.GetProperties(batch, connectionName);
            Assert.True(preBefore.BackgroundQuery, "Precondition: BackgroundQuery should be true before refresh.");

            // Act — refresh must complete without a CPU spin
            _commands.Refresh(batch, connectionName);

            // Assert — BackgroundQuery must be restored to its original value (true)
            var propsAfter = _commands.GetProperties(batch, connectionName);
            Assert.True(propsAfter.BackgroundQuery,
                "BackgroundQuery must be restored to true after refresh. " +
                "If this is false, the fix broke the save/restore logic.");
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
    /// Regression test: BackgroundQuery=false connections must also be preserved.
    /// The fix forces false during refresh; this verifies that a connection that
    /// starts with false is NOT accidentally changed to true after refresh.
    /// </summary>
    [Fact]
    public void Refresh_BackgroundQueryFalse_RemainsAfterRefresh()
    {
        var (testFile, sourceWorkbook, connectionName) = SetupAceOleDbConnection();

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);
            _commands.LoadTo(batch, connectionName, "ProductsData");

            // Change BackgroundQuery to false before the test
            _commands.SetProperties(batch, connectionName, backgroundQuery: false);

            var propsBefore = _commands.GetProperties(batch, connectionName);
            Assert.False(propsBefore.BackgroundQuery, "Precondition: BackgroundQuery should be false.");

            // Act — refresh
            _commands.Refresh(batch, connectionName);

            // Assert — BackgroundQuery must remain false
            var propsAfter = _commands.GetProperties(batch, connectionName);
            Assert.False(propsAfter.BackgroundQuery,
                "BackgroundQuery must remain false after refresh. " +
                "If this is true, the fix accidentally set BackgroundQuery to true.");
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

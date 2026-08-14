using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests
{
    [Fact]
    public void RefreshControl_IdleOleDbConnection_ReportsStatusAndDoesNotCancel()
    {
        var (testFile, sourceWorkbook, connectionName) = SetupAceOleDbConnection();

        try
        {
            using var batch = ExcelSession.BeginBatch(testFile);

            var status = _commands.GetRefreshStatus(batch, connectionName);
            Assert.True(status.Success);
            Assert.True(status.SupportsRefreshStatus);
            Assert.False(status.IsRefreshing);

            var cancel = _commands.CancelRefresh(batch, connectionName);
            Assert.True(cancel.Success);
            Assert.True(cancel.SupportsCancellation);
            Assert.False(cancel.WasRefreshing);
            Assert.False(cancel.Cancelled);
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

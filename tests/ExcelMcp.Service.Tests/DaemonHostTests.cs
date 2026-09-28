using Sbroenne.ExcelMcp.Service.Rpc;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class DaemonHostTests
{
    [Fact(Timeout = 15000)]
    public async Task Shutdown_ResponseArrivesBeforeHostDrainsOpenConnection()
    {
        var pipeName = $"excelmcp-daemon-host-{Guid.NewGuid():N}";
        DaemonHost? host = null;
        host = new DaemonHost(
            request =>
            {
                Assert.Equal("service.shutdown", request.Command);
                host.RequestShutdownAfterResponse();
                return Task.FromResult(new ServiceResponse { Success = true });
            },
            static () => 0);
        using (host)
        {
            var runTask = host.RunAsync(pipeName);
            using var pipe = ServiceSecurity.CreateClient(pipeName);
            await pipe.ConnectAsync(5000);
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            try
            {
                var response = await proxy.ProcessCommandAsync(
                    new ServiceRequest { Command = "service.shutdown" });

                Assert.True(response.Success);
                await Task.Delay(TimeSpan.FromMilliseconds(200));
                Assert.False(
                    runTask.IsCompleted,
                    "The host must drain the retained client connection after sending the shutdown response.");
            }
            finally
            {
                ((IDisposable)proxy).Dispose();
            }

            await runTask.WaitAsync(TimeSpan.FromSeconds(5));
        }
    }
}

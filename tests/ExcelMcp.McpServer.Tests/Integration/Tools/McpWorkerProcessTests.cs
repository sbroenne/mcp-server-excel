using System.Diagnostics;
using System.Runtime.InteropServices;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("Sequential")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceBridge")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Slow")]
public sealed class McpWorkerProcessTests
{
    [Theory(Timeout = 60000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ControllerExit_ClosesProtocolPipesWhileItsBlockedWorkerRemainsAlive(bool useDotnet)
    {
        var executable = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            @"..\..\..\..\..\src\ExcelMcp.McpServer\bin\Release\net10.0-windows\Sbroenne.ExcelMcp.McpServer.exe"));
        var start = new ProcessStartInfo(useDotnet ? "dotnet" : executable)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        if (useDotnet)
            start.ArgumentList.Add(Path.ChangeExtension(executable, ".dll"));
        start.Environment["EXCELMCP_TELEMETRY_MODE"] = "off";
        using var server = Process.Start(start)!;
        var stderr = server.StandardError.ReadToEndAsync();
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        McpClient? client = null;
        Process? worker = null;
        var suspended = false;
        Exception? failure = null;
        try
        {
            client = await McpClient.CreateAsync(new StreamClientTransport(
                serverInput: server.StandardInput.BaseStream,
                serverOutput: server.StandardOutput.BaseStream),
                clientOptions: new McpClientOptions
                {
                    ClientInfo = new() { Name = "worker-process-boundary", Version = "1.0.0" },
                    InitializationTimeout = TimeSpan.FromSeconds(15)
                }, cancellationToken: deadline.Token);
            var result = await client.CallToolAsync("worksheet_read", new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["workbook_session_id"] = "missing-session"
            }, cancellationToken: deadline.Token);
            Assert.True(result.IsError);
            Assert.Contains("not found", McpResponseAssertions.ReadText(result, expectedError: true),
                StringComparison.OrdinalIgnoreCase);
            var registration = Assert.Single(new WorkerRegistry().Read(),
                record => record.Controller == WorkerProcess.From(server));
            Assert.Null(registration.Excel);
            worker = Process.GetProcessById(registration.Worker.ProcessId);
            Assert.Equal(registration.Worker, WorkerProcess.From(worker));
            Assert.Equal(0, NtSuspendProcess(worker.Handle));
            suspended = true;
            server.Kill();
            await server.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(3));
            Assert.False(worker.HasExited);
            var stdout = server.StandardOutput.ReadToEndAsync();
            Assert.Empty(await stdout.WaitAsync(TimeSpan.FromSeconds(3)));
            _ = await stderr.WaitAsync(TimeSpan.FromSeconds(3));
            Assert.False(worker.HasExited);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            try
            {
                if (suspended)
                    Assert.Equal(0, NtResumeProcess(worker!.Handle));
                server.StandardInput.Close();
                await server.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                if (worker is not null)
                    await worker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                if (client is not null)
                    await client.DisposeAsync();
                _ = await stderr.WaitAsync(TimeSpan.FromSeconds(3));
            }
            catch (Exception cleanup)
            {
                failure = failure is null ? cleanup : new AggregateException(failure, cleanup);
            }
            finally
            {
                worker?.Dispose();
            }
        }
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    [DllImport("ntdll.dll")]
    private static extern int NtSuspendProcess(IntPtr processHandle);

    [DllImport("ntdll.dll")]
    private static extern int NtResumeProcess(IntPtr processHandle);
}

using System.IO.Pipelines;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

internal static class ProgramTransportTestHost
{
    private static readonly TimeSpan ClientInitializationTimeout = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan ServerReadyTimeout = TimeSpan.FromSeconds(15);
    private static readonly TimeSpan ServerReadyRetryDelay = TimeSpan.FromMilliseconds(50);
    private static readonly TimeSpan ServerShutdownTimeout =
        ComInteropConstants.StaThreadJoinTimeout + TimeSpan.FromSeconds(15);

    public static async Task<(McpClient Client, Task ServerTask)> StartAsync(
        Pipe clientToServerPipe,
        Pipe serverToClientPipe,
        CancellationToken cancellationToken,
        string clientName,
        Func<ServiceBridge.IServiceBridgeBackend>? serviceFactory = null)
    {
        var serverTask = Program.RunAsync([], clientToServerPipe, serverToClientPipe, cancellationToken, serviceFactory);
        var client = await ConnectClientWithRetryAsync(clientToServerPipe, serverToClientPipe, cancellationToken, clientName);

        return (client, serverTask);
    }

    public static async Task StopAsync(
        McpClient? client,
        Pipe clientToServerPipe,
        Pipe serverToClientPipe,
        Task? serverTask,
        ITestOutputHelper output,
        CancellationTokenSource? cancellationTokenSource = null)
    {
        var failures = new List<Exception>();
        if (client != null)
        {
            try
            {
                await client.DisposeAsync();
            }
            catch (Exception ex)
            {
                failures.Add(new InvalidOperationException("Failed to dispose MCP client.", ex));
            }
        }

        if (serverTask == null)
        {
            await TryCompleteAsync(clientToServerPipe.Writer, failures, nameof(clientToServerPipe) + ".Writer");
            await TryCompleteAsync(serverToClientPipe.Reader, failures, nameof(serverToClientPipe) + ".Reader");
            await TryCompleteAsync(clientToServerPipe.Reader, failures, nameof(clientToServerPipe) + ".Reader");
            await TryCompleteAsync(serverToClientPipe.Writer, failures, nameof(serverToClientPipe) + ".Writer");
            ThrowCleanupFailures(failures);
            return;
        }

        if (cancellationTokenSource is not null)
            await TryCancelAsync(cancellationTokenSource, failures);
        await TryCompleteAsync(clientToServerPipe.Writer, failures, nameof(clientToServerPipe) + ".Writer");
        await TryCompleteAsync(serverToClientPipe.Reader, failures, nameof(serverToClientPipe) + ".Reader");

        try
        {
            await serverTask.WaitAsync(ServerShutdownTimeout);
        }
        catch (OperationCanceledException) when (serverTask.IsCanceled)
        {
        }
        catch (TimeoutException ex)
        {
            failures.Add(ex);
            output.WriteLine("MCP test host exceeded its shutdown deadline; attempting final cleanup.");

            if (cancellationTokenSource is not null && !cancellationTokenSource.IsCancellationRequested)
            {
                await TryCancelAsync(cancellationTokenSource, failures);
            }

            try
            {
                await TryCompleteAsync(clientToServerPipe.Reader, failures, nameof(clientToServerPipe) + ".Reader");
                await TryCompleteAsync(serverToClientPipe.Writer, failures, nameof(serverToClientPipe) + ".Writer");
                await serverTask.WaitAsync(ServerShutdownTimeout);
            }
            catch (OperationCanceledException) when (serverTask.IsCanceled)
            {
            }
            catch (Exception cleanupFailure)
            {
                failures.Add(cleanupFailure);
            }
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }

        await TryCompleteAsync(clientToServerPipe.Reader, failures, nameof(clientToServerPipe) + ".Reader");
        await TryCompleteAsync(serverToClientPipe.Writer, failures, nameof(serverToClientPipe) + ".Writer");

        if (!serverTask.IsCompleted)
        {
            failures.Add(new TimeoutException("MCP test host did not stop after shutdown, forced cancellation, and pipe completion."));
        }
        ThrowCleanupFailures(failures);
    }

    private static void ThrowCleanupFailures(List<Exception> failures)
    {
        if (failures.Count != 0)
            throw new AggregateException("MCP test host cleanup failed.", failures);
    }

    private static async Task TryCancelAsync(CancellationTokenSource source, List<Exception> failures)
    {
        try
        {
            await source.CancelAsync();
        }
        catch (Exception ex)
        {
            failures.Add(new InvalidOperationException("Failed to cancel MCP test host.", ex));
        }
    }

    private static async Task TryCompleteAsync(PipeWriter writer, List<Exception> failures, string pipeName)
    {
        try
        {
            await writer.CompleteAsync();
        }
        catch (Exception ex)
        {
            failures.Add(new InvalidOperationException($"Failed to complete {pipeName}.", ex));
        }
    }

    private static async Task TryCompleteAsync(PipeReader reader, List<Exception> failures, string pipeName)
    {
        try
        {
            await reader.CompleteAsync();
        }
        catch (Exception ex)
        {
            failures.Add(new InvalidOperationException($"Failed to complete {pipeName}.", ex));
        }
    }

    private static async Task<McpClient> ConnectClientWithRetryAsync(
        Pipe clientToServerPipe,
        Pipe serverToClientPipe,
        CancellationToken cancellationToken,
        string clientName)
    {
        var deadline = DateTime.UtcNow + ServerReadyTimeout;

        while (true)
        {
            try
            {
                return await McpClient.CreateAsync(
                    new StreamClientTransport(
                        serverInput: clientToServerPipe.Writer.AsStream(),
                        serverOutput: serverToClientPipe.Reader.AsStream()),
                    clientOptions: new McpClientOptions
                    {
                        ClientInfo = new() { Name = clientName, Version = "1.0.0" },
                        InitializationTimeout = ClientInitializationTimeout
                    },
                    cancellationToken: cancellationToken);
            }
            catch (Exception) when (DateTime.UtcNow < deadline && !cancellationToken.IsCancellationRequested)
            {
                await Task.Delay(ServerReadyRetryDelay, cancellationToken);
            }
        }
    }
}

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
        string clientName)
    {
        Program.ConfigureTestTransport(clientToServerPipe, serverToClientPipe);

        var serverTask = Program.Main([]);
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
        if (client != null)
        {
            try
            {
                await client.DisposeAsync();
            }
            catch (Exception ex)
            {
                output.WriteLine($"Warning: Failed to dispose MCP client cleanly: {ex.Message}");
            }
        }

        if (serverTask == null)
        {
            await TryCompleteAsync(clientToServerPipe.Writer, output, nameof(clientToServerPipe) + ".Writer");
            await TryCompleteAsync(serverToClientPipe.Reader, output, nameof(serverToClientPipe) + ".Reader");
            await TryCompleteAsync(clientToServerPipe.Reader, output, nameof(clientToServerPipe) + ".Reader");
            await TryCompleteAsync(serverToClientPipe.Writer, output, nameof(serverToClientPipe) + ".Writer");
            Program.ResetTestTransport();
            return;
        }

        Program.RequestTestTransportShutdown();
        await TryCompleteAsync(clientToServerPipe.Writer, output, nameof(clientToServerPipe) + ".Writer");
        await TryCompleteAsync(serverToClientPipe.Reader, output, nameof(serverToClientPipe) + ".Reader");

        try
        {
            await serverTask.WaitAsync(ServerShutdownTimeout);
        }
        catch (OperationCanceledException)
        {
        }
        catch (TimeoutException)
        {
            output.WriteLine("Warning: MCP test host did not stop within timeout; forcing cancellation.");

            if (cancellationTokenSource is not null && !cancellationTokenSource.IsCancellationRequested)
            {
                await cancellationTokenSource.CancelAsync();
            }

            try
            {
                await TryCompleteAsync(clientToServerPipe.Reader, output, nameof(clientToServerPipe) + ".Reader");
                await TryCompleteAsync(serverToClientPipe.Writer, output, nameof(serverToClientPipe) + ".Writer");
                await serverTask.WaitAsync(ServerShutdownTimeout);
            }
            catch (OperationCanceledException)
            {
            }
            catch (TimeoutException)
            {
                output.WriteLine("Warning: MCP test host still did not stop after forced cancellation.");
            }
        }
        catch (Exception ex)
        {
            output.WriteLine($"Warning: MCP test host faulted during shutdown: {ex.Message}");
        }

        await TryCompleteAsync(clientToServerPipe.Reader, output, nameof(clientToServerPipe) + ".Reader");
        await TryCompleteAsync(serverToClientPipe.Writer, output, nameof(serverToClientPipe) + ".Writer");

        if (!serverTask.IsCompleted)
        {
            throw new TimeoutException("MCP test host did not stop after shutdown, forced cancellation, and pipe completion.");
        }

        Program.ResetTestTransport();
    }

    private static async Task TryCompleteAsync(PipeWriter writer, ITestOutputHelper output, string pipeName)
    {
        try
        {
            await writer.CompleteAsync();
        }
        catch (Exception ex)
        {
            output.WriteLine($"Warning: Failed to complete {pipeName}: {ex.Message}");
        }
    }

    private static async Task TryCompleteAsync(PipeReader reader, ITestOutputHelper output, string pipeName)
    {
        try
        {
            await reader.CompleteAsync();
        }
        catch (Exception ex)
        {
            output.WriteLine($"Warning: Failed to complete {pipeName}: {ex.Message}");
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

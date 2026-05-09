using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Regression coverage for issue #640.
/// Proves the CLI forwards the caller-supplied datamodel refresh timeout end-to-end.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DataModelRefreshTimeoutRegressionTests
{
    [Fact]
    public async Task Refresh_NumericTimeoutSeconds_ForwardsRequestedBudget()
    {
        var pipeName = $"excelmcp-cli-test-{Guid.NewGuid():N}";

        await using var fakeDaemon = new EchoDaemon(pipeName);
        await fakeDaemon.StartAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["datamodel", "refresh", "--session", "session-issue-640", "--timeout", "600"],
            timeoutMs: 20000,
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            diagnosticLabel: "issue-640-datamodel-refresh-timeout");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("datamodel.refresh", json.RootElement.GetProperty("command").GetString());

        using var argsJson = JsonDocument.Parse(json.RootElement.GetProperty("argsJson").GetString()!);
        Assert.Equal("00:10:00", argsJson.RootElement.GetProperty("timeout").GetString());
    }

    [Fact]
    public async Task Refresh_TimeSpanTimeout_ForwardsRequestedBudget()
    {
        var pipeName = $"excelmcp-cli-test-{Guid.NewGuid():N}";

        await using var fakeDaemon = new EchoDaemon(pipeName);
        await fakeDaemon.StartAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["datamodel", "refresh", "--session", "session-issue-640", "--timeout", "00:10:00"],
            timeoutMs: 20000,
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            diagnosticLabel: "issue-640-datamodel-refresh-timespan-timeout");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());

        using var argsJson = JsonDocument.Parse(json.RootElement.GetProperty("argsJson").GetString()!);
        Assert.Equal("00:10:00", argsJson.RootElement.GetProperty("timeout").GetString());
    }

    [Fact]
    [Trait("Feature", "PowerQuery")]
    public async Task PowerQueryRefresh_NumericTimeoutSeconds_ForwardsRequestedBudget()
    {
        var pipeName = $"excelmcp-cli-test-{Guid.NewGuid():N}";

        await using var fakeDaemon = new EchoDaemon(pipeName);
        await fakeDaemon.StartAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "refresh", "--session", "session-issue-640", "--query-name", "Issue640Query", "--timeout", "600"],
            timeoutMs: 20000,
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            diagnosticLabel: "issue-640-powerquery-refresh-timeout");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("powerquery.refresh", json.RootElement.GetProperty("command").GetString());

        using var argsJson = JsonDocument.Parse(json.RootElement.GetProperty("argsJson").GetString()!);
        Assert.Equal("00:10:00", argsJson.RootElement.GetProperty("timeout").GetString());
    }

    private sealed class EchoDaemon : IAsyncDisposable
    {
        private readonly string _pipeName;
        private readonly CancellationTokenSource _shutdown = new();
        private readonly TaskCompletionSource _listening = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly Task _runTask;

        public EchoDaemon(string pipeName)
        {
            _pipeName = pipeName;
            _runTask = RunAsync();
        }

        public Task StartAsync() => _listening.Task;

        public async ValueTask DisposeAsync()
        {
            _shutdown.Cancel();

            try
            {
                await _runTask;
            }
            catch (OperationCanceledException)
            {
            }

            _shutdown.Dispose();
        }

        private async Task RunAsync()
        {
            _listening.TrySetResult();

            while (!_shutdown.IsCancellationRequested)
            {
                using var server = ServiceSecurity.CreateSecureServer(_pipeName);
                await server.WaitForConnectionAsync(_shutdown.Token);

                var target = new EchoRpcTarget();
                using var rpc = JsonRpc.Attach(server, target);

                try
                {
                    await rpc.Completion.WaitAsync(_shutdown.Token);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
            }
        }
    }

    private sealed class EchoRpcTarget
    {
        private readonly string _pingCommand = "service.ping";

        public Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request)
        {
            if (string.Equals(request.Command, _pingCommand, StringComparison.Ordinal))
            {
                return Task.FromResult(new ServiceResponse
                {
                    Success = true,
                    Command = request.Command,
                    Result = JsonSerializer.Serialize(new
                    {
                        success = true,
                        action = "ping",
                        message = "pong"
                    }, ServiceProtocol.JsonOptions)
                });
            }

            return Task.FromResult(new ServiceResponse
            {
                Success = true,
                Command = request.Command,
                SessionId = request.SessionId,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    command = request.Command,
                    sessionId = request.SessionId,
                    argsJson = request.Args
                }, ServiceProtocol.JsonOptions)
            });
        }
    }
}

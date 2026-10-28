using System.IO.Pipes;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Rpc;
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
[Trait("AdapterTestKind", "System")]
public sealed class DataModelRefreshTimeoutRegressionTests
{
    [Fact]
    public async Task Refresh_NumericTimeoutSeconds_ForwardsRequestedBudget()
    {
        var pipeName = $"excelmcp-cli-test-{Guid.NewGuid():N}";

        await using var fakeDaemon = new EchoDaemon(pipeName);
        await fakeDaemon.StartAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["datamodel", "refresh", "--session", "session-issue-640", "--timeout-seconds", "600"],
            timeoutMs: 20000,
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            diagnosticLabel: "issue-640-datamodel-refresh-timeout");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("datamodel.refresh", json.RootElement.GetProperty("command").GetString());

        using var argsJson = JsonDocument.Parse(json.RootElement.GetProperty("data").GetProperty("argsJson").GetString()!);
        Assert.Equal(600, argsJson.RootElement.GetProperty("timeoutSeconds").GetInt32());
    }

    [Fact]
    public async Task Refresh_TimeSpanTimeout_IsRejectedBeforeDaemonDispatch()
    {
        var pipeName = $"excelmcp-cli-test-{Guid.NewGuid():N}";

        await using var fakeDaemon = new EchoDaemon(pipeName);
        await fakeDaemon.StartAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["datamodel", "refresh", "--session", "session-issue-640", "--timeout-seconds", "00:10:00"],
            timeoutMs: 20000,
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            diagnosticLabel: "issue-640-datamodel-refresh-timespan-timeout");

        Assert.NotEqual(0, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        var error = json.RootElement.GetProperty("error").GetProperty("message").GetString();
        Assert.Contains("00:10:00", error, StringComparison.Ordinal);
        Assert.False(json.RootElement.TryGetProperty("errorMessage", out _));
    }

    [Fact]
    [Trait("Feature", "PowerQuery")]
    public async Task PowerQueryRefresh_NumericTimeoutSeconds_ForwardsRequestedBudget()
    {
        var pipeName = $"excelmcp-cli-test-{Guid.NewGuid():N}";

        await using var fakeDaemon = new EchoDaemon(pipeName);
        await fakeDaemon.StartAsync();

        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["powerquery", "refresh", "--session", "session-issue-640", "--query-name", "Issue640Query", "--timeout-seconds", "600"],
            timeoutMs: 20000,
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = pipeName
            },
            diagnosticLabel: "issue-640-powerquery-refresh-timeout");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("powerquery.refresh", json.RootElement.GetProperty("command").GetString());

        using var argsJson = JsonDocument.Parse(json.RootElement.GetProperty("data").GetProperty("argsJson").GetString()!);
        Assert.Equal(600, argsJson.RootElement.GetProperty("timeoutSeconds").GetInt32());
    }

    private sealed class EchoDaemon : IAsyncDisposable
    {
        private readonly string _pipeName;
        private readonly DaemonHost _work;
        private readonly DaemonHost _control;
        private readonly Task _runTask;
        private readonly TaskCompletionSource _workListening = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource _controlListening = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public EchoDaemon(string pipeName)
        {
            _pipeName = pipeName;
            var target = new EchoRpcTarget();
            _work = new DaemonHost(target.ProcessCommandAsync,
                serverFactory: name => CreateServer(name, _workListening));
            _control = new DaemonHost(target.ProcessCommandAsync,
                serverFactory: name => CreateServer(name, _controlListening));
            _runTask = Task.WhenAll(_work.RunAsync(pipeName),
                _control.RunAsync(ServiceClient.GetControlPipeName(pipeName)));
        }

        public async Task StartAsync()
        {
            await Task.WhenAll(_workListening.Task, _controlListening.Task)
                .WaitAsync(ServiceClient.DefaultConnectTimeout);
            using var client = new ServiceClient(_pipeName, useControlEndpoint: true);
            var response = await client.SendAsync(new ServiceRequest { Command = "service.ping" });
            Assert.True(response.Success,
                $"The test controller did not become ready: {response.ErrorCategory}: {response.ErrorMessage} ({response.ExceptionType}).");
        }

        private static NamedPipeServerStream CreateServer(string pipeName, TaskCompletionSource listening)
        {
            var server = ServiceSecurity.CreateSecureServer(pipeName);
            listening.TrySetResult();
            return server;
        }

        public async ValueTask DisposeAsync()
        {
            _work.RequestShutdown();
            _control.RequestShutdown();
            await _runTask.WaitAsync(TimeSpan.FromSeconds(5));
            _work.Dispose();
            _control.Dispose();
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

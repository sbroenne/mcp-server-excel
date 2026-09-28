using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Tests.Helpers;

internal static class InProcessCliHelper
{
    private static readonly string[] QuietArgument = ["--quiet"];

    internal static async Task<CliResult> RunAsync(
        IReadOnlyList<string> args,
        Func<ServiceRequest, ServiceResponse>? responder = null,
        string input = "")
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var runtime = new CliCommandRuntime(
            new StubClientFactory(responder),
            new StringReader(input),
            output,
            error,
            isOutputRedirected: true);
        var commandArgs = QuietArgument.Concat(args).ToArray();

        var exitCode = await Program.RunAsync(commandArgs, runtime);

        return new CliResult
        {
            ExitCode = exitCode,
            Stdout = output.ToString(),
            Stderr = error.ToString()
        };
    }

    internal static async Task<(CliResult Result, JsonDocument Json)> RunJsonAsync(
        IReadOnlyList<string> args,
        Func<ServiceRequest, ServiceResponse>? responder = null,
        string input = "")
    {
        var result = await RunAsync(args, responder, input);
        return (result, JsonDocument.Parse(result.Stdout));
    }

    internal static async Task<CliResult> RunWithServiceAsync(
        IReadOnlyList<string> args,
        string input = "")
    {
        using var service = new ExcelMcpService();
        return await RunAsync(
            args,
            request => service.ProcessAsync(request).GetAwaiter().GetResult(),
            input);
    }

    internal static Task<CliResult> RunWithServiceAsync(
        string arguments,
        string input = "") =>
        RunWithServiceAsync(SplitArguments(arguments), input);

    internal static async Task<(CliResult Result, JsonDocument Json)> RunJsonWithServiceAsync(
        IReadOnlyList<string> args,
        string input = "")
    {
        var result = await RunWithServiceAsync(args, input);
        return (result, JsonDocument.Parse(result.Stdout));
    }

    internal static async Task<(CliResult Result, IReadOnlyList<ServiceRequest> Requests)>
        RunRecordingAsync(
            string arguments,
            params ServiceResponse[] responses)
    {
        var factory = new RecordingClientFactory(responses);
        var output = new StringWriter();
        var error = new StringWriter();
        var runtime = new CliCommandRuntime(
            factory,
            new StringReader(string.Empty),
            output,
            error,
            isOutputRedirected: true);
        var commandArgs = QuietArgument.Concat(SplitArguments(arguments)).ToArray();

        var exitCode = await Program.RunAsync(commandArgs, runtime);

        return (
            new CliResult
            {
                ExitCode = exitCode,
                Stdout = output.ToString(),
                Stderr = error.ToString()
            },
            factory.Requests);
    }

    private sealed class StubClientFactory(Func<ServiceRequest, ServiceResponse>? responder)
        : ICliRequestClientFactory
    {
        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken)
        {
            if (responder is null)
            {
                throw new InvalidOperationException(
                    "The command unexpectedly attempted Service dispatch.");
            }

            return Task.FromResult<ICliRequestClient>(new StubClient(responder));
        }
    }

    private static List<string> SplitArguments(string arguments)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        var quoted = false;

        for (var index = 0; index < arguments.Length; index++)
        {
            var character = arguments[index];
            if (character == '\\'
                && index + 1 < arguments.Length
                && arguments[index + 1] == '"')
            {
                current.Append('"');
                index++;
                continue;
            }

            if (character == '"')
            {
                quoted = !quoted;
                continue;
            }

            if (char.IsWhiteSpace(character) && !quoted)
            {
                if (current.Length > 0)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                continue;
            }

            current.Append(character);
        }

        if (quoted)
        {
            throw new ArgumentException("Unterminated quoted command argument.", nameof(arguments));
        }

        if (current.Length > 0)
        {
            result.Add(current.ToString());
        }

        return result;
    }

    private sealed class StubClient(Func<ServiceRequest, ServiceResponse> responder)
        : ICliRequestClient
    {
        public Task<ServiceResponse> SendAsync(
            ServiceRequest request,
            CancellationToken cancellationToken) =>
            Task.FromResult(responder(request));

        public void Dispose()
        {
        }
    }

    private sealed class RecordingClientFactory(IEnumerable<ServiceResponse> responses)
        : ICliRequestClientFactory
    {
        private readonly Queue<ServiceResponse> _responses = new(responses);

        internal List<ServiceRequest> Requests { get; } = [];

        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken) =>
            Task.FromResult<ICliRequestClient>(new RecordingClient(this));

        private sealed class RecordingClient(RecordingClientFactory owner) : ICliRequestClient
        {
            public Task<ServiceResponse> SendAsync(
                ServiceRequest request,
                CancellationToken cancellationToken)
            {
                owner.Requests.Add(request);
                return Task.FromResult(owner._responses.Dequeue());
            }

            public void Dispose()
            {
            }
        }
    }
}

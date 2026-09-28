using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal interface ICliRequestClient : IDisposable
{
    Task<ServiceResponse> SendAsync(
        ServiceRequest request,
        CancellationToken cancellationToken);
}

internal interface ICliRequestClientFactory
{
    Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken);
}

internal sealed class CliCommandRuntime
{
    private static readonly AsyncLocal<CliCommandRuntime?> AmbientRuntime = new();
    private static readonly ICliRequestClientFactory ProductionClientFactory =
        new DaemonCliRequestClientFactory();

    internal CliCommandRuntime(
        ICliRequestClientFactory clientFactory,
        TextReader input,
        TextWriter output,
        TextWriter error,
        bool isOutputRedirected)
    {
        ArgumentNullException.ThrowIfNull(clientFactory);
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);
        ClientFactory = clientFactory;
        Input = input;
        Output = output;
        Error = error;
        IsOutputRedirected = isOutputRedirected;
    }

    internal static CliCommandRuntime Current =>
        AmbientRuntime.Value ?? new CliCommandRuntime(
            ProductionClientFactory,
            Console.In,
            Console.Out,
            Console.Error,
            Console.IsOutputRedirected);

    internal ICliRequestClientFactory ClientFactory { get; }
    internal TextReader Input { get; }
    internal TextWriter Output { get; }
    internal TextWriter Error { get; }
    internal bool IsOutputRedirected { get; }

    internal static IDisposable Push(CliCommandRuntime runtime)
    {
        ArgumentNullException.ThrowIfNull(runtime);
        var previous = AmbientRuntime.Value;
        AmbientRuntime.Value = runtime;
        return new RuntimeScope(previous);
    }

    private sealed class RuntimeScope(CliCommandRuntime? previous) : IDisposable
    {
        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            AmbientRuntime.Value = previous;
        }
    }
}

internal sealed class DaemonCliRequestClientFactory : ICliRequestClientFactory
{
    public async Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken)
    {
        var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        return new ServiceClientAdapter(client);
    }

    private sealed class ServiceClientAdapter(ServiceClient client) : ICliRequestClient
    {
        public Task<ServiceResponse> SendAsync(
            ServiceRequest request,
            CancellationToken cancellationToken) =>
            client.SendAsync(request, cancellationToken);

        public void Dispose() => client.Dispose();
    }
}

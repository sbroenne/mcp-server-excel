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

internal interface ICliDaemonConnection
{
    DaemonConnectionPolicy.DaemonObservation Observe(string pipeName);

    Task<ServiceResponse> SendControlRequestAsync(
        string pipeName,
        ServiceRequest request,
        CancellationToken cancellationToken,
        TimeSpan timeout);

    DaemonConnectionPolicy.DaemonFailureState ResolveFailureState(
        string pipeName,
        ServiceResponse response);
}

internal sealed class CliCommandRuntime
{
    private static readonly AsyncLocal<CliCommandRuntime?> AmbientRuntime = new();
    private static readonly ICliRequestClientFactory ProductionClientFactory =
        new DaemonCliRequestClientFactory();
    private static readonly ICliDaemonConnection ProductionDaemonConnection =
        new DaemonConnectionAdapter();

    internal CliCommandRuntime(
        ICliRequestClientFactory clientFactory,
        TextReader input,
        TextWriter output,
        TextWriter error,
        bool isOutputRedirected,
        ICliDaemonConnection? daemonConnection = null,
        Func<Task<string?>>? latestVersionProvider = null)
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
        DaemonConnection = daemonConnection ?? ProductionDaemonConnection;
        LatestVersionProvider = latestVersionProvider
            ?? (() => NuGetVersionChecker.GetLatestVersionAsync());
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
    internal ICliDaemonConnection DaemonConnection { get; }
    internal Func<Task<string?>> LatestVersionProvider { get; }

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

internal sealed class DaemonConnectionAdapter : ICliDaemonConnection
{
    public DaemonConnectionPolicy.DaemonObservation Observe(string pipeName) =>
        DaemonConnectionPolicy.Observe(pipeName);

    public Task<ServiceResponse> SendControlRequestAsync(
        string pipeName,
        ServiceRequest request,
        CancellationToken cancellationToken,
        TimeSpan timeout) =>
        DaemonConnectionPolicy.SendControlRequestAsync(
            pipeName,
            request,
            cancellationToken,
            timeout);

    public DaemonConnectionPolicy.DaemonFailureState ResolveFailureState(
        string pipeName,
        ServiceResponse response) =>
        DaemonConnectionPolicy.ResolveFailureState(pipeName, response);
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

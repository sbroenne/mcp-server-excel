using System.Diagnostics;
using System.Reflection;
using Azure.Monitor.OpenTelemetry.Exporter;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OpenTelemetry;
using OpenTelemetry.Exporter;
using OpenTelemetry.Logs;
using OpenTelemetry.Resources;
using OpenTelemetry.Trace;
using Sbroenne.ExcelMcp.Telemetry;

#if EXCELCLI
namespace Sbroenne.ExcelMcp.CLI.Telemetry;
#else
namespace Sbroenne.ExcelMcp.McpServer.Telemetry;
#endif

// Product threads only publish to an already-built SDK batch processor. No host owns
// these providers: process exit deliberately abandons pending reports without disposal.
internal sealed class UsageTelemetry(
    string serviceName,
    Func<TelemetryUserIdentity>? identityFactory = null,
    Func<TelemetrySettings, TelemetryUserIdentity, UsagePipeline>? pipelineFactory = null)
{
    private readonly string _sessionId = Guid.NewGuid().ToString("N")[..8];
    private UsagePipeline? _pipeline;
    private int _started;
    private int _exiting;
    internal Task Initialization { get; private set; } = Task.CompletedTask;
    internal bool IsReady => Volatile.Read(ref _exiting) == 0 && Volatile.Read(ref _pipeline) != null;

    internal void Start(TelemetrySettings settings)
    {
        if (settings.Mode == TelemetryMode.Off || Volatile.Read(ref _exiting) != 0
            || Interlocked.Exchange(ref _started, 1) != 0)
        {
            return;
        }

        Initialization = Task.Run(() =>
        {
            UsagePipeline? pipeline = null;
            try
            {
                var identity = (identityFactory ?? TelemetryUserIdentityStore.GetOrCreate)();
                if (Volatile.Read(ref _exiting) != 0) return;
                pipeline = pipelineFactory != null
                    ? pipelineFactory(settings, identity)
                    : new UsagePipeline(settings, serviceName, identity, _sessionId);
                if (Volatile.Read(ref _exiting) != 0)
                {
                    pipeline.Dispose(); // Background setup only; never the exit thread.
                    return;
                }
                Volatile.Write(ref _pipeline, pipeline);
                if (TelemetryUserIdentityStore.CreateMigrationEvent(identity) is { } migration)
                {
                    RecordEvent(migration);
                }
                RecordEvent(new UsageEvent("SessionStart"));
            }
            catch (Exception)
            {
                // Reporting failures must not affect operations or write SDK errors to stdout.
                try { pipeline?.Dispose(); }
                catch (Exception) { }
            }
        });
    }

    internal void RecordEvent(UsageEvent telemetry)
    {
        if (!IsReady) return;
        try { Volatile.Read(ref _pipeline)?.RecordEvent(telemetry); }
        catch (Exception) { }
    }

    internal void RecordOperation(UsageOperation telemetry)
    {
        if (!IsReady) return;
        try { Volatile.Read(ref _pipeline)?.RecordOperation(telemetry); }
        catch (Exception) { }
    }

    internal void SignalExit() => Interlocked.Exchange(ref _exiting, 1);
}

internal sealed class UsagePipeline : IDisposable
{
    private readonly ActivitySource _source;
    private readonly TracerProvider _traces;
    private readonly ServiceProvider _services;
    private readonly LoggerProvider _logs;
    private readonly ILogger _logger;
    private readonly TelemetryUserIdentity _identity;
    private readonly string _sessionId;
    private readonly string _version;

    internal UsagePipeline(TelemetrySettings settings, string serviceName,
        TelemetryUserIdentity identity, string sessionId,
        Action<TracerProviderBuilder>? configureTraces = null,
        Action<LoggerProviderBuilder>? configureLogs = null)
    {
        _identity = identity;
        _sessionId = sessionId;
        _version = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "1.0.0";
        var resource = ResourceBuilder.CreateEmpty().AddService(serviceName,
            serviceVersion: _version, serviceInstanceId: identity.RoleInstance);
        _source = new ActivitySource($"{serviceName}.Usage.{Guid.NewGuid():N}");
        var traces = Sdk.CreateTracerProviderBuilder().SetResourceBuilder(resource)
            .AddSource(_source.Name).SetSampler(new AlwaysOnSampler());
        if (configureTraces != null) configureTraces(traces);
        else if (settings.Mode == TelemetryMode.Azure) traces.AddAzureMonitorTraceExporter(o => ConfigureAzure(o, settings));
        else traces.AddOtlpExporter(o => ConfigureLocal(o, settings, "traces"));
        TracerProvider? traceProvider = null;
        ServiceProvider? services = null;
        try
        {
            _traces = traceProvider = traces.Build();
            var collection = new ServiceCollection();
            collection.AddLogging(logging => logging.ClearProviders());
            collection.AddOpenTelemetry().WithLogging(logging =>
            {
                logging.SetResourceBuilder(resource);
                if (configureLogs != null) configureLogs(logging);
                else if (settings.Mode == TelemetryMode.Azure) logging.AddAzureMonitorLogExporter(o => ConfigureAzure(o, settings));
                else logging.AddOtlpExporter(o => ConfigureLocal(o, settings, "logs"));
            });
            _services = services = collection.BuildServiceProvider();
            _logs = _services.GetRequiredService<LoggerProvider>();
            _logger = _services.GetRequiredService<ILoggerFactory>().CreateLogger($"{serviceName}.Usage");
        }
        catch
        {
            services?.Dispose();
            traceProvider?.Dispose();
            _source.Dispose();
            throw;
        }
    }

    internal static void ConfigureAzure(AzureMonitorExporterOptions options, TelemetrySettings settings)
    {
        // The exporter exposes these switches only as environment variables.
        Environment.SetEnvironmentVariable("APPLICATIONINSIGHTS_STATSBEAT_DISABLED", "true");
        Environment.SetEnvironmentVariable("APPLICATIONINSIGHTS_SDKSTATS_DISABLED", "true");
        options.ConnectionString = settings.ConnectionString;
        options.SamplingRatio = 1;
        options.TracesPerSecond = null;
        options.DisableOfflineStorage = true;
        options.EnableLiveMetrics = false;
        options.EnableStandardMetrics = false;
        options.EnablePerformanceCounters = false;
    }

    private static void ConfigureLocal(OtlpExporterOptions options, TelemetrySettings settings, string signal)
    {
        options.Protocol = settings.Protocol == "grpc" ? OtlpExportProtocol.Grpc : OtlpExportProtocol.HttpProtobuf;
        options.Endpoint = options.Protocol == OtlpExportProtocol.Grpc ? settings.Endpoint!
            : new Uri($"{settings.Endpoint!.AbsoluteUri.TrimEnd('/')}/v1/{signal}");
        options.TimeoutMilliseconds = 1000;
        options.Headers = string.Empty;
        options.HttpClientFactory = () => new HttpClient(new SocketsHttpHandler
        {
            UseProxy = false,
            AllowAutoRedirect = false
        })
        {
            Timeout = TimeSpan.FromSeconds(1)
        };
    }

    internal void RecordEvent(UsageEvent telemetry)
    {
        var attributes = Attributes(telemetry.Properties);
        attributes.Add(new("microsoft.custom_event.name", telemetry.Name));
        _logger.Log(LogLevel.Information, new EventId(0, telemetry.Name), attributes, null, (_, _) => telemetry.Name);
    }

    internal void RecordOperation(UsageOperation telemetry)
    {
        var tags = Attributes(telemetry.Properties);
        using var activity = _source.StartActivity(telemetry.Name, ActivityKind.Server,
            default(ActivityContext), tags, startTime: telemetry.Timestamp);
        if (activity == null) return;
        activity.SetStatus(telemetry.Success ? ActivityStatusCode.Ok : ActivityStatusCode.Error);
        activity.SetEndTime((telemetry.Timestamp + telemetry.Duration).UtcDateTime);
    }

    private List<KeyValuePair<string, object?>> Attributes(IDictionary<string, string> properties)
    {
        var attributes = properties.Select(p => new KeyValuePair<string, object?>(p.Key, p.Value)).ToList();
        attributes.Add(new("UserId", _identity.UserId));
        attributes.Add(new("SessionId", _sessionId));
        attributes.Add(new("AppVersion", _version));
        return attributes;
    }

    // Explicit draining is only for deterministic tests, never product shutdown.
    internal bool ForceFlush(int timeoutMilliseconds) => _traces.ForceFlush(timeoutMilliseconds) & _logs.ForceFlush(timeoutMilliseconds);

    public void Dispose()
    {
        _services.Dispose();
        _traces.Dispose();
        _source.Dispose();
    }
}

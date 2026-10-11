using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.Reflection;
using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OpenTelemetry;
using OpenTelemetry.Exporter;
using OpenTelemetry.Logs;
using OpenTelemetry.Resources;
using OpenTelemetry.Trace;

#if EXCELCLI
namespace Sbroenne.ExcelMcp.CLI.Telemetry;
#else
namespace Sbroenne.ExcelMcp.McpServer.Telemetry;
#endif

// Uses already-sanitized telemetry objects, never the request, workbook, or raw exception.
// A separate local pipeline avoids constructing the Azure SDK and its automatic collectors.
internal sealed class LocalTelemetry : IDisposable
{
    private readonly ActivitySource _source;
    private readonly TracerProvider _traces;
    private readonly ServiceProvider _services;
    private readonly LoggerProvider _logs;
    private readonly ILogger _logger;
    private readonly ExporterDiagnostics _diagnostics = new();

    internal LocalTelemetry(
        TelemetrySettings settings,
        string serviceName,
        Action<TracerProviderBuilder>? configureTraces = null,
        Action<LoggerProviderBuilder>? configureLogs = null)
    {
        var version = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        var resource = ResourceBuilder.CreateEmpty().AddService(serviceName, serviceVersion: version);
        _source = new ActivitySource($"{serviceName}.Usage.{Guid.NewGuid():N}");
        var traces = Sdk.CreateTracerProviderBuilder()
            .SetResourceBuilder(resource)
            .AddSource(_source.Name)
            .SetSampler(new AlwaysOnSampler());
        if (configureTraces == null)
        {
            traces.AddOtlpExporter(options => ConfigureExporter(options, settings, "traces"));
        }
        else
        {
            configureTraces(traces);
        }

        TracerProvider? traceProvider = null;
        ServiceProvider? serviceProvider = null;
        try
        {
            _traces = traceProvider = traces.Build();
            var services = new ServiceCollection();
            services.AddLogging(logging => logging.ClearProviders());
            services.AddOpenTelemetry().WithLogging(options =>
            {
                options.SetResourceBuilder(resource);
                if (configureLogs == null)
                {
                    options.AddOtlpExporter(exporter => ConfigureExporter(exporter, settings, "logs"));
                }
                else
                {
                    configureLogs(options);
                }
            });
            _services = serviceProvider = services.BuildServiceProvider();
            _logs = _services.GetRequiredService<LoggerProvider>();
            _logger = _services.GetRequiredService<ILoggerFactory>().CreateLogger($"{serviceName}.Usage");
        }
        catch
        {
            serviceProvider?.Dispose();
            traceProvider?.Dispose();
            _source.Dispose();
            _diagnostics.Dispose();
            throw;
        }
    }

    private static void ConfigureExporter(OtlpExporterOptions options, TelemetrySettings settings, string signal)
    {
        options.Protocol = settings.Protocol == "grpc" ? OtlpExportProtocol.Grpc : OtlpExportProtocol.HttpProtobuf;
        options.Endpoint = options.Protocol == OtlpExportProtocol.Grpc
            ? settings.Endpoint!
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

    internal void Track(EventTelemetry telemetry) =>
        WriteLog(LogLevel.Information, telemetry.Name, telemetry, telemetry.Properties);

    internal void Track(RequestTelemetry telemetry)
    {
        var tags = CreateAttributes(telemetry, telemetry.Properties);
        tags.Add(new("ResponseCode", telemetry.ResponseCode));
        using var activity = _source.StartActivity(
            telemetry.Name, ActivityKind.Server, default(ActivityContext), tags, startTime: telemetry.Timestamp);
        if (activity != null)
        {
            activity.SetStatus(telemetry.Success == false ? ActivityStatusCode.Error : ActivityStatusCode.Ok);
            activity.SetEndTime((telemetry.Timestamp + telemetry.Duration).UtcDateTime);
        }
    }

    internal void Track(ExceptionTelemetry telemetry)
    {
        if (!telemetry.Properties.TryGetValue("Sanitized", out var sanitized) || sanitized != "true")
        {
            throw new InvalidOperationException("Local exception telemetry must be sanitized before recording.");
        }

        WriteLog(LogLevel.Critical, "UnhandledException", telemetry, telemetry.Properties);
    }

    private void WriteLog(LogLevel level, string name, ITelemetry telemetry, IDictionary<string, string> properties)
    {
        var attributes = CreateAttributes(telemetry, properties);
        _logger.Log(level, new EventId(0, name), attributes, null, (_, _) => name);
    }

    private static List<KeyValuePair<string, object?>> CreateAttributes(
        ITelemetry telemetry, IDictionary<string, string> properties)
    {
        var attributes = properties.Select(pair => new KeyValuePair<string, object?>(pair.Key, pair.Value)).ToList();
        AddContext("UserId", telemetry.Context.User.Id);
        AddContext("SessionId", telemetry.Context.Session.Id);
        AddContext("AppVersion", telemetry.Context.Component.Version);
        return attributes;

        void AddContext(string key, string? value)
        {
            if (!attributes.Any(attribute => attribute.Key == key))
            {
                attributes.Add(new(key, value));
            }
        }
    }

    internal bool Flush()
    {
        var tracesFlushed = _traces.ForceFlush(1000);
        var logsFlushed = _logs.ForceFlush(1000);
        if (!tracesFlushed || !logsFlushed)
        {
            Console.Error.WriteLine("Local telemetry could not be delivered. Check that the local OTLP receiver is running.");
        }
        return tracesFlushed && logsFlushed && !_diagnostics.Failed;
    }

    public void Dispose()
    {
        // Flush is bounded separately. Do not add the SDK's default shutdown waits
        // when the local receiver is absent or an export is still in flight.
        _logs.Shutdown(0);
        _traces.Shutdown(0);
        _services.Dispose();
        _traces.Dispose();
        _source.Dispose();
        _diagnostics.Dispose();
    }

    private sealed class ExporterDiagnostics : EventListener
    {
        private int _failed;
        internal bool Failed => Volatile.Read(ref _failed) != 0;

        protected override void OnEventSourceCreated(EventSource eventSource)
        {
            if (eventSource.Name == "OpenTelemetry-Exporter-OpenTelemetryProtocol")
            {
                EnableEvents(eventSource, EventLevel.Warning);
            }
        }

        protected override void OnEventWritten(EventWrittenEventArgs eventData)
        {
            if (Interlocked.Exchange(ref _failed, 1) == 0)
            {
                // Never print the SDK payload: it may contain endpoints, headers, or raw errors.
                Console.Error.WriteLine("Local telemetry delivery failed. Check the local OTLP receiver and protocol settings.");
            }
        }
    }
}

using System.Diagnostics;
using Azure.Monitor.OpenTelemetry.Exporter;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.Extensions.Logging;
using OpenTelemetry;
using OpenTelemetry.Logs;
using OpenTelemetry.Resources;
using OpenTelemetry.Trace;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

/// <summary>
/// Sends CLI command telemetry to Application Insights through OpenTelemetry and the
/// Azure Monitor exporter. It produces the same event and request records the
/// Application Insights <c>TelemetryClient</c> produced, without that client's
/// blocking Azure VM metadata lookup at construction. Each instance owns its own SDK,
/// so nothing is shared through process-wide state.
/// </summary>
internal sealed class OpenTelemetryTelemetrySink : ICliTelemetrySink
{
    // The OpenTelemetry standard switch that turns the SDK off.
    private const string SdkDisabledVariable = "OTEL_SDK_DISABLED";

    // Names the Application Insights SDK used. The category is stored with every event
    // as its CategoryName property, so it must not change.
    private const string EventCategory = "Microsoft.ApplicationInsights.TelemetryClient";
    private const string ActivitySourceName = "Sbroenne.ExcelMcp.CLI.Telemetry";

    // Attribute names the Azure Monitor exporter reads to fill Application Insights fields.
    private const string CustomEventNameAttribute = "microsoft.custom_event.name";
    private const string RequestNameAttribute = "microsoft.request.name";
    private const string RequestResultCodeAttribute = "microsoft.request.resultCode";
    private const string UserIdAttribute = "enduser.pseudo.id";
    private const string SessionIdAttribute = "microsoft.session.id";
    private const string OriginalFormatAttribute = "{OriginalFormat}";

    // Callers bound how long they wait for a flush; this only stops one from hanging forever.
    private const int FlushTimeoutMilliseconds = 10_000;

    // Statsbeat is SDK health data sent to Microsoft, and SDK stats are delivery
    // counters. Neither is CLI usage telemetry, and off Azure each one makes its own
    // blocking 2-second lookup of the Azure VM metadata service.
    private static readonly string[] SdkSelfTelemetryVariables =
    [
        "APPLICATIONINSIGHTS_STATSBEAT_DISABLED",
        "APPLICATIONINSIGHTS_SDKSTATS_DISABLED"
    ];

    private readonly OpenTelemetrySdk? _sdk;
    private readonly ActivitySource? _activitySource;
    private readonly ILogger? _eventLogger;
    private int _disposed;

    private OpenTelemetryTelemetrySink(
        OpenTelemetrySdk? sdk,
        ActivitySource? activitySource,
        ILogger? eventLogger)
    {
        _sdk = sdk;
        _activitySource = activitySource;
        _eventLogger = eventLogger;
    }

    /// <summary>
    /// Builds a sink that sends to <paramref name="connectionString"/>. When the standard
    /// <c>OTEL_SDK_DISABLED</c> variable is <c>true</c> the sink sends nothing.
    /// </summary>
    internal static OpenTelemetryTelemetrySink Create(
        string connectionString,
        TelemetryIdentity identity,
        bool disableOfflineStorage = false)
    {
        if (IsSdkDisabled(Environment.GetEnvironmentVariable))
        {
            return new OpenTelemetryTelemetrySink(null, null, null);
        }

        DisableSdkSelfTelemetry(Environment.GetEnvironmentVariable, Environment.SetEnvironmentVariable);

        var activitySource = new ActivitySource(ActivitySourceName);
        try
        {
            // The exporters are registered per signal. UseAzureMonitorExporter would add
            // them from a hosted service, which only a generic host starts, so with a
            // bare SDK it would send nothing. Metrics need no registration: the exporter
            // creates the standard and performance metrics it reports on its own.
            var sdk = OpenTelemetrySdk.Create(builder =>
            {
                builder.ConfigureResource(resource => ConfigureResource(resource, identity));
                builder.WithTracing(tracing => tracing
                    .AddSource(ActivitySourceName)
                    .AddAzureMonitorTraceExporter(
                        options => ConfigureExporter(options, connectionString, disableOfflineStorage)));
                builder.WithLogging(logging => logging
                    .AddAzureMonitorLogExporter(
                        options => ConfigureExporter(options, connectionString, disableOfflineStorage)));
            });
            return new OpenTelemetryTelemetrySink(
                sdk,
                activitySource,
                sdk.GetLoggerFactory().CreateLogger(EventCategory));
        }
        catch (Exception)
        {
            activitySource.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Sets the role name, instance and version Application Insights shows for every
    /// record. The instance id is the given one; the SDK must not invent its own.
    /// </summary>
    internal static ResourceBuilder ConfigureResource(ResourceBuilder builder, TelemetryIdentity identity) =>
        builder.AddService(
            serviceName: identity.RoleName,
            serviceVersion: identity.Version,
            autoGenerateServiceInstanceId: false,
            serviceInstanceId: identity.RoleInstance);

    /// <summary>
    /// Turns off the SDK's statsbeat and SDK stats features. The exporter reads these
    /// variables when it is created, so they must be set first. A value the user
    /// already set is respected.
    /// </summary>
    internal static void DisableSdkSelfTelemetry(
        Func<string, string?> getVariable,
        Action<string, string> setVariable)
    {
        foreach (var variable in SdkSelfTelemetryVariables)
        {
            if (string.IsNullOrEmpty(getVariable(variable)))
            {
                setVariable(variable, "true");
            }
        }
    }

    /// <summary>Whether the standard <c>OTEL_SDK_DISABLED</c> variable turns telemetry off.</summary>
    internal static bool IsSdkDisabled(Func<string, string?> getVariable) =>
        string.Equals(getVariable(SdkDisabledVariable), "true", StringComparison.OrdinalIgnoreCase);

    public void Track(EventTelemetry eventTelemetry, RequestTelemetry requestTelemetry)
    {
        if (_sdk == null || Volatile.Read(ref _disposed) != 0)
        {
            return;
        }

        // Both records must be roots: a caller's current Activity would otherwise become
        // the parent of the request and add operation ids to the event.
        var ambientActivity = Activity.Current;
        Activity.Current = null;
        try
        {
            TrackEvent(eventTelemetry);
            TrackRequest(requestTelemetry);
        }
        finally
        {
            Activity.Current = ambientActivity;
        }
    }

    public Task FlushAsync() =>
        _sdk == null || Volatile.Read(ref _disposed) != 0
            ? Task.CompletedTask
            : Task.Run(Flush);

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
        {
            return;
        }

        try
        {
            _activitySource?.Dispose();
        }
        catch (Exception)
        {
        }

        try
        {
            // Also exports whatever is left, including the standard and performance metrics.
            _sdk?.Dispose();
        }
        catch (Exception)
        {
        }
    }

    private static void ConfigureExporter(
        AzureMonitorExporterOptions options,
        string connectionString,
        bool disableOfflineStorage)
    {
        options.ConnectionString = connectionString;

        // The default sampler lets through only a share of the spans it sees in the first
        // 200 ms after it is built, and a quick CLI command reports well inside that, so
        // its request record would often be dropped. A limit this high keeps every span
        // and still stamps each request with a microsoft.sample_rate of 100, as before.
        options.TracesPerSecond = 1_000_000;

        // Live Metrics is never used for a CLI, and it kept a connection to the
        // Live Metrics service open for the life of the process.
        options.EnableLiveMetrics = false;

        if (disableOfflineStorage)
        {
            options.DisableOfflineStorage = true;
        }
    }

    // A custom event is a log record carrying the event name. The state mirrors what the
    // Application Insights SDK logged: the item's properties, then the attributes the
    // exporter turns into the event name and tags (it removes them from the properties).
    private void TrackEvent(EventTelemetry? eventTelemetry)
    {
        if (eventTelemetry == null || _eventLogger == null)
        {
            return;
        }

        try
        {
            var state = new List<KeyValuePair<string, object?>>();
            foreach (var property in eventTelemetry.Properties)
            {
                state.Add(new(property.Key, property.Value));
            }

            state.Add(new(CustomEventNameAttribute, eventTelemetry.Name));
            if (!string.IsNullOrEmpty(eventTelemetry.Context.User.Id))
            {
                state.Add(new(UserIdAttribute, eventTelemetry.Context.User.Id));
            }

            if (!string.IsNullOrEmpty(eventTelemetry.Context.Session.Id))
            {
                state.Add(new(SessionIdAttribute, eventTelemetry.Context.Session.Id));
            }

            state.Add(new(OriginalFormatAttribute, string.Empty));

            _eventLogger.Log(LogLevel.Information, 0, state, null, static (_, _) => string.Empty);
        }
        catch (Exception)
        {
        }
    }

    // A request is a server Activity whose start, end and status are set explicitly, so the
    // record carries the command's own timing instead of the time it was tracked.
    private void TrackRequest(RequestTelemetry? requestTelemetry)
    {
        if (requestTelemetry == null || _activitySource == null)
        {
            return;
        }

        try
        {
            using var activity = _activitySource.StartActivity(requestTelemetry.Name, ActivityKind.Server);
            if (activity == null)
            {
                return;
            }

            activity.SetStartTime(requestTelemetry.Timestamp.UtcDateTime);
            activity.SetEndTime(requestTelemetry.Timestamp.Add(requestTelemetry.Duration).UtcDateTime);
            activity.SetStatus(requestTelemetry.Success == true ? ActivityStatusCode.Ok : ActivityStatusCode.Error);

            if (!string.IsNullOrEmpty(requestTelemetry.Name))
            {
                activity.SetTag(RequestNameAttribute, requestTelemetry.Name);
            }

            if (!string.IsNullOrEmpty(requestTelemetry.ResponseCode))
            {
                activity.SetTag(RequestResultCodeAttribute, requestTelemetry.ResponseCode);
            }

            if (!string.IsNullOrEmpty(requestTelemetry.Context.User.Id))
            {
                activity.SetTag(UserIdAttribute, requestTelemetry.Context.User.Id);
            }

            if (!string.IsNullOrEmpty(requestTelemetry.Context.Session.Id))
            {
                activity.SetTag(SessionIdAttribute, requestTelemetry.Context.Session.Id);
            }

            foreach (var property in requestTelemetry.Properties)
            {
                activity.SetTag(property.Key, property.Value);
            }
        }
        catch (Exception)
        {
        }
    }

    private void Flush()
    {
        try
        {
            _sdk?.TracerProvider.ForceFlush(FlushTimeoutMilliseconds);
        }
        catch (Exception)
        {
        }

        try
        {
            _sdk?.LoggerProvider.ForceFlush(FlushTimeoutMilliseconds);
        }
        catch (Exception)
        {
        }
    }
}

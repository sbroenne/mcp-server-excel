using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

internal static class CliTelemetry
{
    private const string EntryPoint = "cli";
    private static readonly string SessionId = Guid.NewGuid().ToString("N")[..8];
    private static readonly string UserId = GenerateAnonymousUserId();
    private static TelemetryClient? _telemetryClient;
    private static TelemetryConfiguration? _telemetryConfiguration;

    internal static void Initialize()
    {
        var connectionString = GetConnectionString();
        if (connectionString == null)
        {
            return;
        }

        try
        {
            _telemetryConfiguration = TelemetryConfiguration.CreateDefault();
            _telemetryConfiguration.ConnectionString = connectionString;
            _telemetryClient = new TelemetryClient(_telemetryConfiguration);
        }
        catch (Exception)
        {
            _telemetryClient = null;
            _telemetryConfiguration?.Dispose();
            _telemetryConfiguration = null;
        }
    }

    internal static async Task<ServiceResponse> TrackCommandAsync(
        ServiceRequest request,
        Func<Task<ServiceResponse>> operation) =>
        await TrackCommandAsync(request, operation, TrackCommandInvocation);

    internal static async Task<ServiceResponse> TrackCommandAsync(
        ServiceRequest request,
        Func<Task<ServiceResponse>> operation,
        Action<string, long, bool, string?> trackInvocation)
    {
        var stopwatch = Stopwatch.StartNew();
        ServiceResponse? response = null;
        try
        {
            response = await operation();
            return response;
        }
        finally
        {
            stopwatch.Stop();
            trackInvocation(
                request.Command,
                stopwatch.ElapsedMilliseconds,
                response?.Success == true,
                response?.ErrorCategory);
        }
    }

    internal static void Flush()
    {
        if (_telemetryClient == null)
        {
            return;
        }

        try
        {
            _telemetryClient.FlushAsync(CancellationToken.None).Wait(TimeSpan.FromSeconds(2));
        }
        catch (Exception)
        {
        }
        finally
        {
            _telemetryConfiguration?.Dispose();
            _telemetryConfiguration = null;
            _telemetryClient = null;
        }
    }

    private static string? GetConnectionString() =>
        string.IsNullOrEmpty(TelemetryConfig.ConnectionString)
        || TelemetryConfig.ConnectionString.StartsWith("__", StringComparison.Ordinal)
            ? null
            : TelemetryConfig.ConnectionString;

    private static void TrackCommandInvocation(
        string command,
        long durationMs,
        bool succeeded,
        string? errorCategory)
    {
        if (_telemetryClient == null)
        {
            return;
        }

        try
        {
            var (eventTelemetry, requestTelemetry) =
                CreateCommandInvocationTelemetry(command, durationMs, succeeded, errorCategory);
            _telemetryClient.TrackEvent(eventTelemetry);
            _telemetryClient.TrackRequest(requestTelemetry);
        }
        catch (Exception)
        {
        }
    }

    internal static (EventTelemetry Event, RequestTelemetry Request)
        CreateCommandInvocationTelemetry(
            string command,
            long durationMs,
            bool succeeded,
            string? errorCategory)
    {
        var parts = command.Split('.', 2);
        if (parts.Length != 2
            || string.IsNullOrWhiteSpace(parts[0])
            || string.IsNullOrWhiteSpace(parts[1]))
        {
            throw new ArgumentException("Telemetry command must use category.action format.", nameof(command));
        }

        var tool = parts[0];
        var action = parts[1];
        var operationName = $"{tool}/{action}";
        var duration = TimeSpan.FromMilliseconds(durationMs);
        var properties = new Dictionary<string, string>
        {
            ["Tool"] = tool,
            ["Action"] = action,
            ["EntryPoint"] = EntryPoint,
            ["Success"] = succeeded.ToString(),
            ["Outcome"] = succeeded ? "succeeded" : "failed"
        };

        if (!succeeded)
        {
            properties["FailureClass"] = ClassifyFailure(errorCategory);
        }

        var eventTelemetry = new EventTelemetry(operationName);
        foreach (var property in properties)
        {
            eventTelemetry.Properties[property.Key] = property.Value;
        }
        eventTelemetry.Properties["DurationMs"] =
            durationMs.ToString(CultureInfo.InvariantCulture);
        ApplyContext(eventTelemetry);

        var requestTelemetry = new RequestTelemetry
        {
            Name = operationName,
            Timestamp = DateTimeOffset.UtcNow.Subtract(duration),
            Duration = duration,
            ResponseCode = succeeded ? "200" : "500",
            Success = succeeded
        };
        foreach (var property in properties)
        {
            requestTelemetry.Properties[property.Key] = property.Value;
        }
        ApplyContext(requestTelemetry);

        return (eventTelemetry, requestTelemetry);
    }

    private static string ClassifyFailure(string? errorCategory) =>
        errorCategory switch
        {
            "InvalidInput" or "SessionNotFound" or "SessionUnavailable"
                or "NotFound" or "Conflict" or "Syntax" or "Expression" or "Prerequisite"
                => "input-state",
            "Privacy" or "Authentication" or "Connectivity" or "Permissions"
                or "DependencyUnavailable" => "external-dependency",
            "Timeout" or "Cancelled" or "SessionInvalidated" => "timeout-cancellation",
            "ComInterop" or "ExcelProcessDied" or "Cleanup" => "excel-runtime",
            "ServiceUnavailable" or "ServiceStartup" or "InvalidResponse"
                => "internal-product-fault",
            _ => "unclassified"
        };

    private static void ApplyContext(ITelemetry telemetry)
    {
        telemetry.Context.User.Id ??= UserId;
        telemetry.Context.Session.Id ??= SessionId;
        telemetry.Context.Cloud.RoleName ??= "ExcelMcp.CLI";
        telemetry.Context.Cloud.RoleInstance = $"instance-{UserId[..8]}";
        telemetry.Context.Component.Version = GetVersion();
    }

    private static string GetVersion() =>
        typeof(CliTelemetry).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? typeof(CliTelemetry).Assembly.GetName().Version?.ToString()
        ?? "1.0.0";

    private static string GenerateAnonymousUserId()
    {
        try
        {
            var machineIdentity =
                $"{Environment.MachineName}|{Environment.UserName}|{Environment.OSVersion.Platform}";
            return Convert.ToHexString(
                SHA256.HashData(Encoding.UTF8.GetBytes(machineIdentity)))[..16]
                .ToLowerInvariant();
        }
        catch (Exception)
        {
            return Guid.NewGuid().ToString("N")[..16];
        }
    }
}

using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

internal static class CliTelemetry
{
    private const string EntryPoint = "cli";
    internal const string UnknownOperationPart = "other";
    private const string UnknownCommand = $"{UnknownOperationPart}.{UnknownOperationPart}";

    /// <summary>Commands handled by the service host instead of a generated category.</summary>
    private static readonly Dictionary<string, string[]> BuiltInActions = new(StringComparer.OrdinalIgnoreCase)
    {
        ["service"] = ["start", "stop", "status", "shutdown", "ping"],
        ["session"] = ["create", "open", "close", "list", "test"],
        ["batch"] = ["run"]
    };

    private static readonly string[] HelpFlags = ["--help", "-h"];
    private static readonly AsyncLocal<InvocationTelemetryState?> CurrentInvocationTelemetry = new();
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
        string? failureCategory = null;
        try
        {
            response = await operation();
            return response;
        }
        catch (Exception ex)
        {
            // Transport failures never reach a response, so classify the thrown
            // exception with the shared classifier instead of losing the reason.
            failureCategory = OperationFailureClassifier.Classify(ex);
            throw;
        }
        finally
        {
            stopwatch.Stop();
            var invocationTelemetry = CurrentInvocationTelemetry.Value;
            var trackedFailureCategory = response?.ErrorCategory ?? failureCategory;
            if (invocationTelemetry != null && response?.Success == false)
            {
                invocationTelemetry.FailureCategory ??= trackedFailureCategory;
            }

            if (invocationTelemetry?.TrackRequests is not false)
            {
                if (invocationTelemetry != null)
                {
                    invocationTelemetry.RequestTracked = true;
                }
                trackInvocation(
                    request.Command,
                    stopwatch.ElapsedMilliseconds,
                    response?.Success == true,
                    trackedFailureCategory);
            }
        }
    }

    /// <summary>
    /// Tracks the whole CLI invocation so commands that never issue a service
    /// request, and failures raised before one is sent, are still measured.
    /// Regular commands emit their final outcome; batches retain per-item telemetry.
    /// </summary>
    internal static int TrackCliInvocation(string[] args, Func<int> operation) =>
        TrackCliInvocation(args, operation, TrackCommandInvocation);

    internal static int TrackCliInvocation(
        string[] args,
        Func<int> operation,
        Action<string, long, bool, string?> trackInvocation)
    {
        if (args.Any(arg => HelpFlags.Contains(arg, StringComparer.OrdinalIgnoreCase)))
        {
            return operation();
        }

        var isBatch = IsBatchCommand(args);
        var previousInvocationTelemetry = CurrentInvocationTelemetry.Value;
        var invocationTelemetry = new InvocationTelemetryState(isBatch);
        CurrentInvocationTelemetry.Value = invocationTelemetry;
        var stopwatch = Stopwatch.StartNew();
        var exitCode = 1;
        string? failureCategory = null;
        try
        {
            exitCode = operation();
            return exitCode;
        }
        catch (Exception ex)
        {
            failureCategory = OperationFailureClassifier.Classify(ex);
            throw;
        }
        finally
        {
            stopwatch.Stop();
            CurrentInvocationTelemetry.Value = previousInvocationTelemetry;
            if (!isBatch || !invocationTelemetry.RequestTracked)
            {
                trackInvocation(
                    ResolveCliCommand(args),
                    stopwatch.ElapsedMilliseconds,
                    exitCode == 0,
                    failureCategory ?? invocationTelemetry.FailureCategory);
            }
        }
    }

    private static bool IsBatchCommand(string[] args) =>
        args.Length > 0 && string.Equals(args[0], "batch", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Maps parsed CLI arguments to a canonical service command. Arguments that
    /// are not an allowlisted command never leave the machine.
    /// </summary>
    internal static string ResolveCliCommand(string[] args)
    {
        if (args.Length == 0)
        {
            return UnknownCommand;
        }

        var category = ServiceRegistry.CategoryByCliCommand.TryGetValue(args[0], out var mapped)
            ? mapped
            : args[0];
        // Commands without a subcommand (for example "batch") use the "run" action.
        var action = args.Length > 1 && !args[1].StartsWith('-') ? args[1] : "run";
        return $"{category}.{action}";
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
        var (tool, action) = ResolveOperation(command);
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

    /// <summary>
    /// Maps a service command to allowlisted telemetry parts. Commands outside
    /// the generated and built-in command sets report fixed labels so
    /// user-supplied text is never sent.
    /// </summary>
    internal static (string Tool, string Action) ResolveOperation(string? command)
    {
        var parts = command?.Split('.', 2) ?? [];
        if (parts.Length != 2)
        {
            return (UnknownOperationPart, UnknownOperationPart);
        }

        var category = parts[0];
        var action = parts[1];
        IReadOnlyList<string>? validActions =
            BuiltInActions.TryGetValue(category, out var builtIn) ? builtIn
            : ServiceRegistry.ValidActionsByCategory.TryGetValue(category, out var generated) ? generated
            : null;
        if (validActions == null)
        {
            return (UnknownOperationPart, UnknownOperationPart);
        }

        var canonicalAction = validActions
            .FirstOrDefault(valid => string.Equals(valid, action, StringComparison.OrdinalIgnoreCase));
        return canonicalAction == null
            ? (UnknownOperationPart, UnknownOperationPart)
            : (category.ToLowerInvariant(), canonicalAction);
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

    private sealed class InvocationTelemetryState(bool trackRequests)
    {
        public bool TrackRequests { get; } = trackRequests;

        public bool RequestTracked { get; set; }

        public string? FailureCategory { get; set; }
    }
}

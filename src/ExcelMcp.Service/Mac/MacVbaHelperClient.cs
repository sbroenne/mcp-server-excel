using System.Diagnostics;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed record MacVbaHelperInstallation(
    bool IsConfigured,
    bool SourceExists,
    string? HelperPath,
    string Status);

internal sealed class MacVbaHelperException(
    string category,
    string code,
    string message) : InvalidOperationException(message)
{
    public string Category { get; } = category;
    public string Code { get; } = code;
}

internal sealed class MacVbaHelperClient(MacExcelBackend backend)
{
    private const string HelperPathEnvironmentVariable = "EXCELMCP_MAC_VBA_HELPER_PATH";

    public static MacVbaHelperInstallation GetInstallation()
    {
        var configuredPath = Environment.GetEnvironmentVariable(HelperPathEnvironmentVariable);
        if (string.IsNullOrWhiteSpace(configuredPath))
        {
            return new MacVbaHelperInstallation(
                false,
                false,
                null,
                $"Set {HelperPathEnvironmentVariable} to the exact installed ExcelMcpHelper.xlam path.");
        }

        var fullPath = Path.GetFullPath(configuredPath);
        if (!string.Equals(
                Path.GetFileName(fullPath),
                "ExcelMcpHelper.xlam",
                StringComparison.Ordinal))
        {
            return new MacVbaHelperInstallation(
                true,
                false,
                fullPath,
                "The configured helper path must end with ExcelMcpHelper.xlam.");
        }

        return new MacVbaHelperInstallation(
            true,
            File.Exists(fullPath),
            fullPath,
            File.Exists(fullPath)
                ? "The configured helper add-in exists; runtime identity and capabilities are not yet proven."
                : "The configured helper add-in does not exist at the exact configured path.");
    }

    public async Task<JsonElement> DispatchAsync(
        string workbookFullName,
        string action,
        object arguments,
        TimeSpan timeout)
    {
        var startedAt = Stopwatch.GetTimestamp();
        if (!string.Equals(action, "helper.capabilities", StringComparison.Ordinal))
        {
            var capabilities = await GetCapabilitiesAsync(
                workbookFullName,
                RemainingTimeout(startedAt, timeout));
            var supported = capabilities.GetProperty("supportedActions")
                .EnumerateArray()
                .Any(item => string.Equals(item.GetString(), action, StringComparison.Ordinal));
            if (!supported)
            {
                throw new MacVbaHelperException(
                    "PlatformNotSupported",
                    "helper_action_unavailable",
                    "The configured helper does not advertise the requested action.");
            }
        }

        return await DispatchCoreAsync(
            workbookFullName,
            action,
            arguments,
            RemainingTimeout(startedAt, timeout));
    }

    private static TimeSpan RemainingTimeout(long startedAt, TimeSpan timeout)
    {
        if (timeout == Timeout.InfiniteTimeSpan)
        {
            return timeout;
        }

        var remaining = timeout - Stopwatch.GetElapsedTime(startedAt);
        if (remaining <= TimeSpan.Zero)
        {
            throw new TimeoutException(
                "The macOS VBA helper operation exceeded its shared capability-and-dispatch deadline. " +
                "The workbook session is no longer safe to use.");
        }
        return remaining;
    }

    public async Task<JsonElement> GetCapabilitiesAsync(
        string workbookFullName,
        TimeSpan timeout)
    {
        var capabilities = await DispatchCoreAsync(
            workbookFullName,
            "helper.capabilities",
            new { },
            timeout);
        if (!capabilities.TryGetProperty("helperVersion", out var helperVersionElement)
            || helperVersionElement.ValueKind != JsonValueKind.String
            || !capabilities.TryGetProperty("protocolVersion", out var protocolVersionElement)
            || protocolVersionElement.ValueKind != JsonValueKind.Number
            || !capabilities.TryGetProperty("supportedActions", out var supportedActions)
            || supportedActions.ValueKind != JsonValueKind.Array
            || supportedActions.EnumerateArray().Any(item => item.ValueKind != JsonValueKind.String)
            || !capabilities.TryGetProperty("staticAvailability", out var staticAvailability)
            || staticAvailability.ValueKind != JsonValueKind.Object
            || !capabilities.TryGetProperty("engineCapabilities", out var engineCapabilities)
            || engineCapabilities.ValueKind != JsonValueKind.Object
            || !capabilities.TryGetProperty("trustReadiness", out var trustReadiness)
            || trustReadiness.ValueKind != JsonValueKind.Object
            || !capabilities.TryGetProperty("provenMethods", out var provenMethods)
            || provenMethods.ValueKind != JsonValueKind.Object)
        {
            throw new MacVbaHelperException(
                "InvalidInput",
                "helper_capabilities_invalid",
                "The configured helper returned an invalid capability response.");
        }
        var helperVersion = helperVersionElement.GetString();
        var protocolVersion = protocolVersionElement.GetInt32();
        if (!string.Equals(
                helperVersion,
                MacVbaHelperProtocol.HelperVersion,
                StringComparison.Ordinal)
            || protocolVersion != MacVbaHelperProtocol.Version)
        {
            throw new MacVbaHelperException(
                "Conflict",
                "helper_version_mismatch",
                "The configured helper add-in version does not match this ExcelMcp build.");
        }
        return capabilities;
    }

    private async Task<JsonElement> DispatchCoreAsync(
        string workbookFullName,
        string action,
        object arguments,
        TimeSpan timeout)
    {
        var installation = GetInstallation();
        if (!installation.IsConfigured
            || !installation.SourceExists
            || installation.HelperPath is null)
        {
            throw new MacVbaHelperException(
                "NotFound",
                "helper_not_configured",
                installation.Status);
        }

        var requestId = Guid.NewGuid().ToString("N");
        var canonicalWorkbookPath = Path.GetFullPath(workbookFullName);
        var requestJson = MacVbaHelperProtocol.CreateRequest(
            requestId,
            canonicalWorkbookPath,
            action,
            arguments);
        var dispatchResult = await backend.InvokeAsync(
            "helper.dispatch",
            new
            {
                helperPath = installation.HelperPath,
                requestJson
            },
            timeout);
        if (!dispatchResult.TryGetProperty("responseJson", out var responseJsonElement)
            || responseJsonElement.GetString() is not { } responseJson)
        {
            throw new InvalidOperationException(
                "The helper transport returned no correlated JSON response.");
        }

        var response = MacVbaHelperProtocol.ParseResponse(responseJson, requestId);
        if (!response.Success)
        {
            var error = response.Error
                ?? throw new InvalidOperationException("The helper returned an invalid failure response.");
            throw new MacVbaHelperException(error.Category, error.Code, error.Message);
        }

        return response.Result?.Clone()
            ?? JsonSerializer.SerializeToElement(new { }, ServiceProtocol.JsonOptions);
    }
}

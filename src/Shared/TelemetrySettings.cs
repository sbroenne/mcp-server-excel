using System.Net;

#if EXCELCLI
namespace Sbroenne.ExcelMcp.CLI.Telemetry;
#else
namespace Sbroenne.ExcelMcp.McpServer.Telemetry;
#endif

internal enum TelemetryMode
{
    Off,
    Local,
    Azure
}

internal sealed record TelemetrySettings(TelemetryMode Mode, string? ConnectionString, Uri? Endpoint, string? Protocol)
{
    internal const string ModeVariable = "EXCELMCP_TELEMETRY_MODE";

    internal static TelemetrySettings Read(string? embeddedConnectionString) =>
        Resolve(embeddedConnectionString, Environment.GetEnvironmentVariable);

    internal static TelemetrySettings Resolve(string? embeddedConnectionString, Func<string, string?> getVariable)
    {
        var connectionString = string.IsNullOrEmpty(embeddedConnectionString)
            || embeddedConnectionString.StartsWith("__", StringComparison.Ordinal)
                ? null : embeddedConnectionString;
        var value = getVariable(ModeVariable)?.Trim().ToLowerInvariant();
        var mode = value switch
        {
            null or "" => connectionString == null ? TelemetryMode.Off : TelemetryMode.Azure,
            "off" => TelemetryMode.Off,
            "local" => TelemetryMode.Local,
            "azure" => TelemetryMode.Azure,
            _ => throw new InvalidOperationException($"{ModeVariable} must be off, local, or azure.")
        };

        if (mode == TelemetryMode.Azure && connectionString == null)
        {
            throw new InvalidOperationException("Azure telemetry requires a build-time Application Insights connection string.");
        }

        if (mode != TelemetryMode.Local)
        {
            return new(mode, mode == TelemetryMode.Azure ? connectionString : null, null, null);
        }

        var protocol = getVariable("OTEL_EXPORTER_OTLP_PROTOCOL") ?? "grpc";
        if (protocol is not ("grpc" or "http/protobuf"))
        {
            throw new InvalidOperationException("Local telemetry OTEL_EXPORTER_OTLP_PROTOCOL must be grpc or http/protobuf.");
        }

        var endpointValue = getVariable("OTEL_EXPORTER_OTLP_ENDPOINT")
            ?? (protocol == "grpc" ? "http://localhost:4317" : "http://localhost:4318");
        if (!Uri.TryCreate(endpointValue, UriKind.Absolute, out var endpoint)
            || endpoint.Scheme is not ("http" or "https")
            || !(endpoint.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase)
                || (IPAddress.TryParse(endpoint.Host.Trim('[', ']'), out var address) && IPAddress.IsLoopback(address)))
            || !string.IsNullOrEmpty(endpoint.UserInfo)
            || !string.IsNullOrEmpty(endpoint.Query)
            || !string.IsNullOrEmpty(endpoint.Fragment))
        {
            throw new InvalidOperationException("Local telemetry OTEL_EXPORTER_OTLP_ENDPOINT must be an HTTP or HTTPS loopback address without credentials, a query, or a fragment.");
        }

        return new(mode, null, endpoint, protocol);
    }
}

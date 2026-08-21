using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class DaemonTrackingJson
{
    internal static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() }
    };
}

using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Severity assigned to an integrity finding.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WorkbookIntegritySeverity>))]
public enum WorkbookIntegritySeverity
{
    /// <summary>A workbook integrity failure.</summary>
    [JsonStringEnumMemberName("error")]
    Error,

    /// <summary>A condition that needs review but may be intentional or temporary.</summary>
    [JsonStringEnumMemberName("warning")]
    Warning,

    /// <summary>Context that does not fail validation.</summary>
    [JsonStringEnumMemberName("information")]
    Information
}

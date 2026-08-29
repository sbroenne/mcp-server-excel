using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Whether a finding is directly reported by Excel or inferred from a pattern.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WorkbookIntegrityReliability>))]
public enum WorkbookIntegrityReliability
{
    /// <summary>The finding is based on directly observable workbook state.</summary>
    [JsonStringEnumMemberName("deterministic")]
    Deterministic,

    /// <summary>The finding is inferred and may represent intentional workbook design.</summary>
    [JsonStringEnumMemberName("heuristic")]
    Heuristic
}

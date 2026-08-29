using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Overall workbook integrity outcome.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WorkbookIntegrityStatus>))]
public enum WorkbookIntegrityStatus
{
    /// <summary>No errors or warnings were found.</summary>
    [JsonStringEnumMemberName("passed")]
    Passed,

    /// <summary>No errors were found, but one or more warnings need review.</summary>
    [JsonStringEnumMemberName("passed-with-warnings")]
    PassedWithWarnings,

    /// <summary>One or more integrity errors were found.</summary>
    [JsonStringEnumMemberName("failed")]
    Failed
}

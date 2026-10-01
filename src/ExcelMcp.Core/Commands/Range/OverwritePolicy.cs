using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Whether a range content write may replace existing cell content.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<OverwritePolicy>))]
public enum OverwritePolicy
{
    /// <summary>Reject occupied destinations before any write; inspection failures stop the operation.</summary>
    [JsonStringEnumMemberName("reject-nonempty")]
    [EnumAlias("reject-nonempty")]
    RejectNonempty,

    /// <summary>Allow intentional replacement without bypassing Excel's write restrictions.</summary>
    [JsonStringEnumMemberName("allow")]
    Allow
}

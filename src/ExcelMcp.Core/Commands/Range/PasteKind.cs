using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>The native Excel paste content to transfer.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<PasteKind>))]
public enum PasteKind
{
    /// <summary>Values, formulas, formatting, and other native cell properties.</summary>
    [JsonStringEnumMemberName("all")]
    All,
    /// <summary>Calculated values without formulas or source formatting.</summary>
    [JsonStringEnumMemberName("values")]
    Values,
    /// <summary>Formulas with native reference adjustment; constants and blanks also transfer.</summary>
    [JsonStringEnumMemberName("formulas")]
    Formulas,
    /// <summary>Native formatting, including number formats and protection, without content.</summary>
    [JsonStringEnumMemberName("formats")]
    Formats,
    /// <summary>Validation rules without content or visual formatting.</summary>
    [JsonStringEnumMemberName("validation")]
    Validation
}

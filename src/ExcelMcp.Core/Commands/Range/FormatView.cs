using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>The native formatting snapshot to inspect.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<FormatView>))]
public enum FormatView
{
    /// <summary>Workbook cell formatting, without conditional formatting.</summary>
    [JsonStringEnumMemberName("stored")]
    Stored,
    /// <summary>Effective displayed formatting, including conditional formatting.</summary>
    [JsonStringEnumMemberName("displayed")]
    Displayed,
    /// <summary>Both stored and effective displayed snapshots.</summary>
    [JsonStringEnumMemberName("both")]
    Both
}

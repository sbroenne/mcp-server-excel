using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Current or changed global native database query approval setting.
/// </summary>
public sealed class NativeQueryApprovalResult : ResultBase
{
    /// <summary>Whether approval is required before a new native database query runs.</summary>
    public bool Required { get; set; }

    /// <summary>Whether the requirement is explicitly stored instead of using Excel's default.</summary>
    public bool ExplicitlySet { get; set; }

    /// <summary>Path to the current user's Power Query settings archive.</summary>
    public string SettingsFilePath { get; set; } = string.Empty;

    /// <summary>Security warning that this setting applies to every workbook for the current Windows user.</summary>
    public string Warning { get; set; } = string.Empty;

    /// <summary>Previous approval requirement when a setting change was requested.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? PreviousRequired { get; set; }

    /// <summary>Backup archive created before changing the setting.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BackupPath { get; set; }
}

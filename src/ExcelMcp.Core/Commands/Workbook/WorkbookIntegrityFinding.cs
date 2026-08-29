using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>One workbook integrity concern with structured context and remediation.</summary>
public sealed class WorkbookIntegrityFinding
{
    /// <summary>Stable machine-readable finding code.</summary>
    public string Code { get; set; } = string.Empty;

    /// <summary>Finding severity.</summary>
    public WorkbookIntegritySeverity Severity { get; set; }

    /// <summary>Finding category.</summary>
    public WorkbookIntegrityCategory Category { get; set; }

    /// <summary>Whether the finding is directly observed or inferred.</summary>
    public WorkbookIntegrityReliability Reliability { get; set; }

    /// <summary>Human-readable description of the problem.</summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>Suggested corrective action.</summary>
    public string SuggestedRemediation { get; set; } = string.Empty;

    /// <summary>Affected worksheet, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SheetName { get; set; }

    /// <summary>Affected A1 cell address, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CellAddress { get; set; }

    /// <summary>Affected formula text, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula { get; set; }

    /// <summary>Canonical Excel error name, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorName { get; set; }

    /// <summary>Raw Excel COM error code, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ErrorCode { get; set; }

    /// <summary>Affected table name, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TableName { get; set; }

    /// <summary>Affected table column name, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ColumnName { get; set; }

    /// <summary>External workbook link source, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LinkSource { get; set; }

    /// <summary>Normalized Excel link status, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LinkStatus { get; set; }

    /// <summary>Caller-supplied expected control-total value, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? ExpectedValue { get; set; }

    /// <summary>Observed value or canonical error name, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? ActualValue { get; set; }

    /// <summary>Caller-supplied absolute control-total tolerance, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Tolerance { get; set; }
}

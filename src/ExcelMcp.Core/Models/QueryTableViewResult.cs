namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Detailed QueryTable configuration.
/// </summary>
public sealed class QueryTableViewResult : ResultBase
{
    /// <summary>QueryTable name.</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Worksheet containing the QueryTable.</summary>
    public string SheetName { get; init; } = string.Empty;

    /// <summary>Top-left destination address.</summary>
    public string Destination { get; init; } = string.Empty;

    /// <summary>Source type: text, web, database, or other.</summary>
    public string SourceType { get; init; } = string.Empty;

    /// <summary>Connection string reported by Excel.</summary>
    public string Connection { get; init; } = string.Empty;

    /// <summary>Whether refresh executes in the background.</summary>
    public bool BackgroundQuery { get; init; }

    /// <summary>Whether the query refreshes when the workbook opens.</summary>
    public bool RefreshOnFileOpen { get; init; }

    /// <summary>Automatic refresh period in minutes.</summary>
    public int RefreshPeriod { get; init; }

    /// <summary>Whether Excel adjusts column widths after refresh.</summary>
    public bool AdjustColumnWidth { get; init; }

    /// <summary>Whether Excel preserves formatting after refresh.</summary>
    public bool PreserveFormatting { get; init; }

    /// <summary>Text import delimiter, when applicable.</summary>
    public string? Delimiter { get; init; }

    /// <summary>
    /// Excel's reported TextFilePlatform value, when applicable.
    /// Excel can normalize this legacy value after refresh, so it may differ from the requested code page.
    /// </summary>
    public int? Encoding { get; init; }

    /// <summary>Web selection type, when applicable.</summary>
    public string? WebSelectionType { get; init; }

    /// <summary>Selected HTML table indexes, when applicable.</summary>
    public string? WebTables { get; init; }

    /// <summary>Web formatting mode, when applicable.</summary>
    public string? WebFormatting { get; init; }
}

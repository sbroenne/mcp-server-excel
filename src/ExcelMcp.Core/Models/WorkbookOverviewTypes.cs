using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Bounded workbook structure and optional cell preview.</summary>
public sealed class WorkbookOverviewResult : ResultBase
{
    /// <summary>Worksheet structure and visibility, when requested.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WorkbookOverviewSheetSection? Sheets { get; set; }

    /// <summary>Excel tables, when requested.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WorkbookOverviewTableSection? Tables { get; set; }

    /// <summary>Visible user-defined names, when requested.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WorkbookOverviewNameSection? DefinedNames { get; set; }

    /// <summary>Bounded cell preview, when requested.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WorkbookOverviewPreview? Preview { get; set; }
}

/// <summary>Bounded worksheet metadata and omission count.</summary>
public sealed class WorkbookOverviewSheetSection
{
    /// <summary>Number of worksheets in the selected scope.</summary>
    public int Count { get; set; }
    /// <summary>Number of worksheets not returned due to maxItems.</summary>
    public int OmittedCount { get; set; }
    /// <summary>Returned worksheets in workbook order.</summary>
    public List<WorkbookOverviewSheet> Items { get; set; } = [];
}

/// <summary>Worksheet identity, visibility, and used-range dimensions.</summary>
public sealed class WorkbookOverviewSheet
{
    /// <summary>Worksheet name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Visible, Hidden, or VeryHidden.</summary>
    public string Visibility { get; set; } = string.Empty;
    /// <summary>Excel's used-range address.</summary>
    public string UsedRangeAddress { get; set; } = string.Empty;
    /// <summary>Row count of the used range.</summary>
    public int UsedRowCount { get; set; }
    /// <summary>Column count of the used range.</summary>
    public int UsedColumnCount { get; set; }
}

/// <summary>Bounded table metadata and omission count.</summary>
public sealed class WorkbookOverviewTableSection
{
    /// <summary>Number of tables in the selected scope.</summary>
    public int Count { get; set; }
    /// <summary>Number of tables not returned due to maxItems.</summary>
    public int OmittedCount { get; set; }
    /// <summary>Returned tables in workbook order.</summary>
    public List<WorkbookOverviewTable> Items { get; set; } = [];
}

/// <summary>Excel table identity and location.</summary>
public sealed class WorkbookOverviewTable
{
    /// <summary>Table name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Worksheet containing the table.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Excel address of the table.</summary>
    public string RangeAddress { get; set; } = string.Empty;
}

/// <summary>Bounded visible-name metadata and omission count.</summary>
public sealed class WorkbookOverviewNameSection
{
    /// <summary>Number of visible user-defined names.</summary>
    public int Count { get; set; }
    /// <summary>Number of names not returned due to maxItems.</summary>
    public int OmittedCount { get; set; }
    /// <summary>Returned names in Excel order.</summary>
    public List<WorkbookOverviewName> Items { get; set; } = [];
}

/// <summary>Visible user-defined name and its Excel RefersTo expression.</summary>
public sealed class WorkbookOverviewName
{
    /// <summary>Name as reported by Excel.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Excel RefersTo expression; may refer to cells, formulas, or external locations.</summary>
    public string RefersTo { get; set; } = string.Empty;
}

/// <summary>Bounded values and formulas from one selected worksheet range.</summary>
public sealed class WorkbookOverviewPreview
{
    /// <summary>Worksheet containing the preview.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Actual range read by Excel.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Rows returned.</summary>
    public int RowCount { get; set; }
    /// <summary>Columns returned.</summary>
    public int ColumnCount { get; set; }
    /// <summary>Rows in the requested scope that were not read.</summary>
    public int OmittedRowCount { get; set; }
    /// <summary>Columns in the requested scope that were not read.</summary>
    public int OmittedColumnCount { get; set; }
    /// <summary>Text characters returned across values and formulas.</summary>
    public int TextCharactersReturned { get; set; }
    /// <summary>Number of text cells shortened by the per-cell or total character limit.</summary>
    public int TruncatedTextCellCount { get; set; }
    /// <summary>Cell values as a row-major matrix.</summary>
    public List<List<object?>> Values { get; set; } = [];
    /// <summary>Cell formulas/contents as a row-major matrix.</summary>
    public List<List<object?>> Formulas { get; set; } = [];
}

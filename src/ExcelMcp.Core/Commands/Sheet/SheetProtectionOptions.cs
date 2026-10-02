using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>Native protected-sheet selection modes.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<ProtectedSheetSelection>))]
public enum ProtectedSheetSelection
{
    /// <summary>Allow selecting any cell.</summary>
    AllCells,
    /// <summary>Allow selecting only unlocked cells.</summary>
    UnlockedCells,
    /// <summary>Prevent cell selection.</summary>
    None
}

/// <summary>Native Worksheet.Protect options; omitted permissions use Excel's restrictive defaults.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class SheetProtectionOptions
{
    /// <summary>Protect drawing objects.</summary>
    public bool DrawingObjects { get; set; } = true;
    /// <summary>Protect locked cells.</summary>
    public bool Contents { get; set; } = true;
    /// <summary>Protect scenarios.</summary>
    public bool Scenarios { get; set; } = true;
    /// <summary>Protect the UI but allow automation; runtime-only, not retained after reopening.</summary>
    public bool UserInterfaceOnly { get; set; }
    /// <summary>Allow cell formatting.</summary>
    public bool AllowFormattingCells { get; set; }
    /// <summary>Allow column formatting.</summary>
    public bool AllowFormattingColumns { get; set; }
    /// <summary>Allow row formatting.</summary>
    public bool AllowFormattingRows { get; set; }
    /// <summary>Allow inserting columns.</summary>
    public bool AllowInsertingColumns { get; set; }
    /// <summary>Allow inserting rows.</summary>
    public bool AllowInsertingRows { get; set; }
    /// <summary>Allow inserting hyperlinks.</summary>
    public bool AllowInsertingHyperlinks { get; set; }
    /// <summary>Allow deleting columns; Excel still requires unlocked cells.</summary>
    public bool AllowDeletingColumns { get; set; }
    /// <summary>Allow deleting rows; Excel still requires unlocked cells.</summary>
    public bool AllowDeletingRows { get; set; }
    /// <summary>Allow sorting; Excel still requires unlocked cells.</summary>
    public bool AllowSorting { get; set; }
    /// <summary>Allow changing an existing AutoFilter, not creating/removing one.</summary>
    public bool AllowFiltering { get; set; }
    /// <summary>Allow using PivotTables.</summary>
    public bool AllowUsingPivotTables { get; set; }
    /// <summary>Optional cell-selection restriction; omitted leaves the native setting unchanged.</summary>
    public ProtectedSheetSelection? Selection { get; set; }
}

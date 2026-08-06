using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Outline;

/// <summary>Worksheet outline axis.</summary>
public enum OutlineAxis
{
    /// <summary>Group complete worksheet rows.</summary>
    Row,
    /// <summary>Group complete worksheet columns.</summary>
    Column,
}

/// <summary>Explicit row/column outline state.</summary>
public sealed class OutlineStateResult : OperationResult
{
    /// <summary>Worksheet containing the outline.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Normalized axis: row or column.</summary>
    public string Axis { get; set; } = string.Empty;
    /// <summary>Canonical full-row or full-column A1 address.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Uniform public level, or null when the range is mixed.</summary>
    public int? Level { get; set; }
    /// <summary>Minimum public outline level across addressed units.</summary>
    public int MinimumLevel { get; set; }
    /// <summary>Maximum public outline level across addressed units.</summary>
    public int MaximumLevel { get; set; }
    /// <summary>Uniform hidden/collapsed state, or null when mixed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Collapsed { get; set; }
    /// <summary>Number of addressed rows or columns.</summary>
    public int UnitCount { get; set; }
    /// <summary>Whether set-level changed level or collapse state.</summary>
    public bool Changed { get; set; }
}

/// <summary>
/// Deterministic row/column grouping. Level 0 means ungrouped; levels 1-7 are nested groups.
/// Ranges must cover complete rows (for row axis) or complete columns (for column axis).
/// No selection, active-cell inference, or toggle actions are used.
/// </summary>
[ServiceCategory("outline", "Outline")]
[McpTool("outline", Title = "Row and Column Outlines", Destructive = true, Category = "structure",
    Description = "Set or inspect explicit worksheet row/column grouping. level 0 means ungrouped; 1-7 are nested groups. Use complete row ranges such as 5:10 or complete column ranges such as B:D. set-level is idempotent and accepts an explicit collapsed value; no selection or toggle behavior.")]
public interface IOutlineCommands
{
    /// <summary>Sets an exact outline level and optional explicit collapsed state.</summary>
    [ServiceAction("set-level")]
    OutlineStateResult SetLevel(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter] int level,
        [RequiredParameter][FromString] OutlineAxis axis,
        bool? collapsed = null);

    /// <summary>Reads outline level/collapse state for an exact full-row or full-column range.</summary>
    [ServiceAction("get-state", IsMutation = false)]
    OutlineStateResult GetState(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter][FromString] OutlineAxis axis);
}

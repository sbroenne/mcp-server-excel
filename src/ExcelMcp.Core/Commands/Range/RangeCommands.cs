
namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Excel range operations implementation - unified API for all range data operations.
/// Single cell is treated as 1x1 range. Named ranges work transparently via rangeAddress parameter.
/// All operations are COM-backed (no data processing in server).
/// Implements IRangeCommands (values/formulas/copy/clear/discovery),
/// IRangeEditCommands (insert/delete/find/replace/sort),
/// IRangeFormatCommands (styling/validation/merge/autofit),
/// and IRangeLinkCommands (hyperlinks/cell protection).
/// </summary>
public partial class RangeCommands : IRangeCommands, IRangeEditCommands, IRangeFormatCommands, IRangeLinkCommands
{
    // This is the main partial class file containing only the class declaration.
    // Implementation methods are organized into separate partial files by feature domain:
    // - RangeCommands.Values.cs (GetValues, SetValues)
    // - RangeCommands.Formulas.cs (GetFormulas, SetFormulas)
    // - RangeCommands.Editing.cs (Clear, Copy, Insert, Delete operations)
    // - RangeCommands.Search.cs (Find, Replace, Sort)
    // - RangeCommands.Discovery.cs (GetUsedRange, GetCurrentRegion, GetRangeInfo)
    // - RangeCommands.Hyperlinks.cs (Add, Remove, List, Get hyperlinks)
    // - RangeCommands.NumberFormat.cs (Get, Set number formats)
    // - RangeCommands.Formatting.cs (SetStyle, GetStyle, FormatRange)
    // - RangeCommands.Validation.cs (ValidateRange, GetValidation, RemoveValidation)
    // - RangeCommands.AutoFit.cs (AutoFitColumns, AutoFitRows)
    // - RangeCommands.Advanced.cs (MergeCells, UnmergeCells, GetMergeInfo, SetCellLock, GetCellLock)
}



#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Excel range operations implementation - unified API for all range data operations.
/// Single cell is treated as 1x1 range. Named ranges work transparently via rangeAddress parameter.
/// All operations are COM-backed (no data processing in server).
/// </summary>
public partial class RangeCommands : IRangeCommands
{
    // This is the main partial class file containing only the class declaration.
    // Implementation methods are organized into separate partial files by feature domain:
    // - RangeCommands.Values.cs (GetValues, SetValues)
    // - RangeCommands.Formulas.cs (GetFormulas, SetFormulas)
    // - RangeCommands.Editing.cs (Clear, Copy, Insert, Delete operations)
    // - RangeCommands.Search.cs (Find, Replace, Sort)
    // - RangeCommands.Discovery.cs (GetUsedRange, GetCurrentRegion, GetRangeInfo)
    // - RangeCommands.Hyperlinks.cs (Add, Remove, List, Get hyperlinks)
}

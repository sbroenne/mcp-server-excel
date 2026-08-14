using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

/// <summary>
/// Result of an Excel Goal Seek operation.
/// </summary>
public sealed class GoalSeekResult : OperationResult
{
    /// <summary>Whether Excel found a solution.</summary>
    public bool Converged { get; set; }

    /// <summary>Final value in the formula cell.</summary>
    public double FormulaValue { get; set; }

    /// <summary>Final value in the changing cell.</summary>
    public double ChangingValue { get; set; }
}

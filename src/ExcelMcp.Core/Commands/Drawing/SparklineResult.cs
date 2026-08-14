using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>Result containing one sparkline group.</summary>
public sealed class SparklineResult : OperationResult
{
    /// <summary>Sparkline details.</summary>
    public SparklineInfo Sparkline { get; set; } = new();
}

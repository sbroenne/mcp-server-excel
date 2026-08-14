using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>Result containing sparkline groups.</summary>
public sealed class SparklineListResult : OperationResult
{
    /// <summary>Sparkline groups on the worksheet.</summary>
    public List<SparklineInfo> Sparklines { get; set; } = [];
}

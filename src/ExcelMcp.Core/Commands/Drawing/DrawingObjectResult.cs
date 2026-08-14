using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>Result containing one worksheet drawing object.</summary>
public sealed class DrawingObjectResult : OperationResult
{
    /// <summary>Drawing object details.</summary>
    public DrawingObjectInfo DrawingObject { get; set; } = new();
}

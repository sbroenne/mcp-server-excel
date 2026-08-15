using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>Result containing worksheet drawing objects.</summary>
public sealed class DrawingObjectListResult : OperationResult
{
    /// <summary>Drawing objects on the worksheet.</summary>
    public List<DrawingObjectInfo> DrawingObjects { get; set; } = [];
}

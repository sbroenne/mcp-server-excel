namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>Information about a worksheet drawing object.</summary>
public sealed class DrawingObjectInfo
{
    /// <summary>Shape name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Containing worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Drawing object category.</summary>
    public DrawingObjectKind Kind { get; set; }
    /// <summary>AutoShape type, when applicable.</summary>
    public DrawingShapeType? ShapeType { get; set; }
    /// <summary>Connector type, when applicable.</summary>
    public DrawingConnectorType? ConnectorType { get; set; }
    /// <summary>Forms control type, when applicable.</summary>
    public DrawingFormControlType? FormControlType { get; set; }
    /// <summary>Left position in points.</summary>
    public double Left { get; set; }
    /// <summary>Top position in points.</summary>
    public double Top { get; set; }
    /// <summary>Width in points.</summary>
    public double Width { get; set; }
    /// <summary>Height in points.</summary>
    public double Height { get; set; }
    /// <summary>Rotation in degrees.</summary>
    public double Rotation { get; set; }
    /// <summary>Whether the object is visible.</summary>
    public bool Visible { get; set; }
    /// <summary>Whether the object is locked on protected sheets.</summary>
    public bool Locked { get; set; }
    /// <summary>Placement mode: 1 move and size, 2 move only, 3 free floating.</summary>
    public int? Placement { get; set; }
    /// <summary>Displayed text.</summary>
    public string? Text { get; set; }
    /// <summary>Text size in points.</summary>
    public double? FontSize { get; set; }
    /// <summary>Text color as #RRGGBB.</summary>
    public string? FontColor { get; set; }
    /// <summary>Fill color as #RRGGBB.</summary>
    public string? FillColor { get; set; }
    /// <summary>Line color as #RRGGBB.</summary>
    public string? LineColor { get; set; }
    /// <summary>Line weight in points.</summary>
    public double? LineWeight { get; set; }
    /// <summary>Accessibility alternative text.</summary>
    public string? AlternativeText { get; set; }
    /// <summary>Linked cell for a Forms control.</summary>
    public string? LinkedCell { get; set; }
    /// <summary>Input range for list and drop-down Forms controls.</summary>
    public string? InputRange { get; set; }
}

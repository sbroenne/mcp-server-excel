namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

/// <summary>Worksheet drawing object categories.</summary>
public enum DrawingObjectKind
{
    /// <summary>Unclassified worksheet shape.</summary>
    Other = 0,
    /// <summary>Excel AutoShape.</summary>
    AutoShape = 1,
    /// <summary>Embedded or linked picture.</summary>
    Image = 2,
    /// <summary>Text box.</summary>
    TextBox = 3,
    /// <summary>Straight, elbow, or curved connector.</summary>
    Connector = 4,
    /// <summary>Worksheet Forms control.</summary>
    FormControl = 5
}

/// <summary>Supported Excel AutoShape types.</summary>
public enum DrawingShapeType
{
    /// <summary>Rectangle.</summary>
    Rectangle = 1,
    /// <summary>Parallelogram.</summary>
    Parallelogram = 2,
    /// <summary>Trapezoid.</summary>
    Trapezoid = 3,
    /// <summary>Diamond.</summary>
    Diamond = 4,
    /// <summary>Rounded rectangle.</summary>
    RoundedRectangle = 5,
    /// <summary>Octagon.</summary>
    Octagon = 6,
    /// <summary>Isosceles triangle.</summary>
    IsoscelesTriangle = 7,
    /// <summary>Right triangle.</summary>
    RightTriangle = 8,
    /// <summary>Oval.</summary>
    Oval = 9,
    /// <summary>Hexagon.</summary>
    Hexagon = 10,
    /// <summary>Cross.</summary>
    Cross = 11,
    /// <summary>Regular pentagon.</summary>
    RegularPentagon = 12,
    /// <summary>Cylinder/can.</summary>
    Can = 13,
    /// <summary>Cube.</summary>
    Cube = 14,
    /// <summary>Bevel.</summary>
    Bevel = 15,
    /// <summary>Folded corner.</summary>
    FoldedCorner = 16,
    /// <summary>Smiley face.</summary>
    SmileyFace = 17,
    /// <summary>Donut.</summary>
    Donut = 18,
    /// <summary>No-smoking symbol.</summary>
    NoSmoking = 19,
    /// <summary>Block arc.</summary>
    BlockArc = 20,
    /// <summary>Heart.</summary>
    Heart = 21,
    /// <summary>Lightning bolt.</summary>
    LightningBolt = 22,
    /// <summary>Sun.</summary>
    Sun = 23,
    /// <summary>Moon.</summary>
    Moon = 24,
    /// <summary>Arc.</summary>
    Arc = 25,
    /// <summary>Right arrow.</summary>
    RightArrow = 33,
    /// <summary>Left arrow.</summary>
    LeftArrow = 34,
    /// <summary>Up arrow.</summary>
    UpArrow = 35,
    /// <summary>Down arrow.</summary>
    DownArrow = 36,
    /// <summary>Left-right arrow.</summary>
    LeftRightArrow = 37,
    /// <summary>Up-down arrow.</summary>
    UpDownArrow = 38,
    /// <summary>Four-direction arrow.</summary>
    QuadArrow = 39,
    /// <summary>Flowchart process.</summary>
    FlowchartProcess = 61,
    /// <summary>Flowchart decision.</summary>
    FlowchartDecision = 63,
    /// <summary>Flowchart data.</summary>
    FlowchartData = 64,
    /// <summary>Flowchart predefined process.</summary>
    FlowchartPredefinedProcess = 65,
    /// <summary>Flowchart internal storage.</summary>
    FlowchartInternalStorage = 66,
    /// <summary>Flowchart document.</summary>
    FlowchartDocument = 67,
    /// <summary>Flowchart multiple documents.</summary>
    FlowchartMultidocument = 68,
    /// <summary>Flowchart terminator.</summary>
    FlowchartTerminator = 69,
    /// <summary>Flowchart preparation.</summary>
    FlowchartPreparation = 70,
    /// <summary>Flowchart manual input.</summary>
    FlowchartManualInput = 71,
    /// <summary>Flowchart manual operation.</summary>
    FlowchartManualOperation = 72,
    /// <summary>Flowchart connector.</summary>
    FlowchartConnector = 73,
    /// <summary>Flowchart off-page connector.</summary>
    FlowchartOffpageConnector = 74
}

/// <summary>Supported connector geometries.</summary>
public enum DrawingConnectorType
{
    /// <summary>Straight connector.</summary>
    Straight = 1,
    /// <summary>Elbow connector.</summary>
    Elbow = 2,
    /// <summary>Curved connector.</summary>
    Curved = 3
}

/// <summary>
/// Safe worksheet Forms controls. ActiveX/OLE controls are intentionally excluded.
/// </summary>
public enum DrawingFormControlType
{
    /// <summary>Push button.</summary>
    Button = 0,
    /// <summary>Check box.</summary>
    CheckBox = 1,
    /// <summary>Drop-down list.</summary>
    DropDown = 2,
    /// <summary>Group box.</summary>
    GroupBox = 4,
    /// <summary>Static label.</summary>
    Label = 5,
    /// <summary>List box.</summary>
    ListBox = 6,
    /// <summary>Option button.</summary>
    OptionButton = 7,
    /// <summary>Scroll bar.</summary>
    ScrollBar = 8,
    /// <summary>Spinner.</summary>
    Spinner = 9
}

/// <summary>Supported sparkline chart types.</summary>
public enum DrawingSparklineType
{
    /// <summary>Line sparkline.</summary>
    Line = 1,
    /// <summary>Column sparkline.</summary>
    Column = 2,
    /// <summary>Win/loss sparkline.</summary>
    WinLoss = 3
}

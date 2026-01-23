namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Marker style constants for Excel chart series.
/// Excel COM: XlMarkerStyle enumeration.
/// </summary>
public enum MarkerStyle
{
    /// <summary>No marker (xlMarkerStyleNone)</summary>
    None = -4142,

    /// <summary>Automatic marker (xlMarkerStyleAutomatic)</summary>
    Automatic = -4105,

    /// <summary>Circle marker (xlMarkerStyleCircle)</summary>
    Circle = 8,

    /// <summary>Dash marker (xlMarkerStyleDash)</summary>
    Dash = -4115,

    /// <summary>Diamond marker (xlMarkerStyleDiamond)</summary>
    Diamond = 2,

    /// <summary>Dot marker (xlMarkerStyleDot)</summary>
    Dot = -4118,

    /// <summary>Picture marker (xlMarkerStylePicture)</summary>
    Picture = -4147,

    /// <summary>Plus sign marker (xlMarkerStylePlus)</summary>
    Plus = 9,

    /// <summary>Square marker (xlMarkerStyleSquare)</summary>
    Square = 1,

    /// <summary>Star marker (xlMarkerStyleStar)</summary>
    Star = 5,

    /// <summary>Triangle marker (xlMarkerStyleTriangle)</summary>
    Triangle = 3,

    /// <summary>X marker (xlMarkerStyleX)</summary>
    X = -4168
}

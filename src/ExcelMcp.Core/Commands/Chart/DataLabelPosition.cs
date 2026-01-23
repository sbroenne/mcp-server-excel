namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Data label position constants for Excel charts.
/// Excel COM: XlDataLabelPosition enumeration.
/// </summary>
public enum DataLabelPosition
{
    /// <summary>Best fit determined by Excel (xlLabelPositionBestFit)</summary>
    BestFit = 5,

    /// <summary>Center of the data point (xlLabelPositionCenter)</summary>
    Center = -4108,

    /// <summary>Above the data point (xlLabelPositionAbove)</summary>
    Above = 0,

    /// <summary>Below the data point (xlLabelPositionBelow)</summary>
    Below = 1,

    /// <summary>Left of the data point (xlLabelPositionLeft)</summary>
    Left = -4131,

    /// <summary>Right of the data point (xlLabelPositionRight)</summary>
    Right = -4152,

    /// <summary>Inside base of bar/column (xlLabelPositionInsideBase)</summary>
    InsideBase = 4,

    /// <summary>Inside end of bar/column (xlLabelPositionInsideEnd)</summary>
    InsideEnd = 3,

    /// <summary>Outside end of bar/column (xlLabelPositionOutsideEnd)</summary>
    OutsideEnd = 2,

    /// <summary>Mixed positions (read-only)</summary>
    Mixed = 6
}

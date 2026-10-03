using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>Native date timeline display levels.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<TimelineGranularity>))]
public enum TimelineGranularity
{
    /// <summary>Year display.</summary>
    Years,
    /// <summary>Quarter display.</summary>
    Quarters,
    /// <summary>Month display.</summary>
    Months,
    /// <summary>Day display.</summary>
    Days
}

/// <summary>Patch selected native slicer/timeline visual properties.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class SlicerUpdateOptions
{
    /// <summary>Left coordinate in points.</summary>
    public double? Left { get; set; }
    /// <summary>Top coordinate in points.</summary>
    public double? Top { get; set; }
    /// <summary>Positive width in points.</summary>
    public double? Width { get; set; }
    /// <summary>Positive height in points.</summary>
    public double? Height { get; set; }
    /// <summary>Caption text.</summary>
    public string? Caption { get; set; }
    /// <summary>Existing native style name.</summary>
    public string? Style { get; set; }
    /// <summary>Positive ordinary slicer column count; not valid for timelines.</summary>
    public int? ColumnCount { get; set; }
    /// <summary>Ordinary slicer header visibility.</summary>
    public bool? DisplayHeader { get; set; }
    /// <summary>Timeline display level.</summary>
    public TimelineGranularity? Granularity { get; set; }
    /// <summary>Timeline header visibility.</summary>
    public bool? ShowHeader { get; set; }
    /// <summary>Timeline selection label visibility.</summary>
    public bool? ShowSelectionLabel { get; set; }
    /// <summary>Timeline time-level selector visibility.</summary>
    public bool? ShowTimeLevel { get; set; }
    /// <summary>Timeline horizontal scrollbar visibility.</summary>
    public bool? ShowHorizontalScrollbar { get; set; }
}

/// <summary>Inclusive native timeline date selection.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class TimelineSelectionOptions
{
    /// <summary>Inclusive starting calendar date.</summary>
    public required DateTime StartDate { get; set; }
    /// <summary>Inclusive ending calendar date.</summary>
    public required DateTime EndDate { get; set; }
}

/// <summary>Native timeline display and date state.</summary>
public sealed class TimelineDetails
{
    /// <summary>Display granularity.</summary>
    public TimelineGranularity Granularity { get; set; }
    /// <summary>Header visibility.</summary>
    public bool ShowHeader { get; set; }
    /// <summary>Selection-label visibility.</summary>
    public bool ShowSelectionLabel { get; set; }
    /// <summary>Time-level selector visibility.</summary>
    public bool ShowTimeLevel { get; set; }
    /// <summary>Horizontal scrollbar visibility.</summary>
    public bool ShowHorizontalScrollbar { get; set; }
    /// <summary>Native start date, null when date filtering is cleared.</summary>
    public DateTime? StartDate { get; set; }
    /// <summary>Native end date, null when date filtering is cleared.</summary>
    public DateTime? EndDate { get; set; }
    /// <summary>Native filter type, null when date filtering is cleared.</summary>
    public string? FilterType { get; set; }
    /// <summary>Whether the native selection uses a single date range; null when filtering is cleared.</summary>
    public bool? SingleRangeFilterState { get; set; }
}

/// <summary>Complete selected visual-control state.</summary>
public sealed class SlicerDetails
{
    /// <summary>Control name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Native shared cache name.</summary>
    public string CacheName { get; set; } = string.Empty;
    /// <summary>Native source field name.</summary>
    public string FieldName { get; set; } = string.Empty;
    /// <summary>Destination worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Displayed caption.</summary>
    public string Caption { get; set; } = string.Empty;
    /// <summary>Style name.</summary>
    public string Style { get; set; } = string.Empty;
    /// <summary>Left coordinate in points.</summary>
    public double Left { get; set; }
    /// <summary>Top coordinate in points.</summary>
    public double Top { get; set; }
    /// <summary>Width in points.</summary>
    public double Width { get; set; }
    /// <summary>Height in points.</summary>
    public double Height { get; set; }
    /// <summary>Ordinary column count; null for timelines.</summary>
    public int? ColumnCount { get; set; }
    /// <summary>Ordinary header visibility; null for timelines.</summary>
    public bool? DisplayHeader { get; set; }
    /// <summary>Whether this is a timeline.</summary>
    public bool IsTimeline { get; set; }
    /// <summary>Whether this is an Excel Table control.</summary>
    public bool IsTable { get; set; }
    /// <summary>Whether the native shared cache filter is cleared.</summary>
    public bool FilterCleared { get; set; }
    /// <summary>All connected PivotTables.</summary>
    public List<string> ConnectedPivotTables { get; set; } = [];
    /// <summary>Source Excel Table name, where applicable.</summary>
    public string? ConnectedTable { get; set; }
    /// <summary>All available native items; empty for timelines.</summary>
    public List<string> AvailableItems { get; set; } = [];
    /// <summary>All selected native items; empty for timelines.</summary>
    public List<string> SelectedItems { get; set; } = [];
    /// <summary>Native timeline state, where applicable.</summary>
    public TimelineDetails? Timeline { get; set; }
}

/// <summary>Result of inspecting or changing a slicer/timeline.</summary>
public sealed class SlicerStateResult : ResultBase
{
    /// <summary>Complete native visual-control state.</summary>
    public SlicerDetails Slicer { get; set; } = new();
}

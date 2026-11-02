using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>Overrides for a copied slicer style. Omitted elements retain their source formatting.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class SlicerStyleOptions
{
    /// <summary>Entire slicer frame.</summary>
    public SlicerStyleElementOptions? WholeTable { get; set; }
    /// <summary>Slicer header.</summary>
    public SlicerStyleElementOptions? Header { get; set; }
    /// <summary>Selected button with data.</summary>
    public SlicerStyleElementOptions? SelectedItemWithData { get; set; }
    /// <summary>Selected button without data.</summary>
    public SlicerStyleElementOptions? SelectedItemWithNoData { get; set; }
    /// <summary>Unselected button with data.</summary>
    public SlicerStyleElementOptions? UnselectedItemWithData { get; set; }
    /// <summary>Unselected button without data.</summary>
    public SlicerStyleElementOptions? UnselectedItemWithNoData { get; set; }
    /// <summary>Hovered selected button with data.</summary>
    public SlicerStyleElementOptions? HoveredSelectedItemWithData { get; set; }
    /// <summary>Hovered selected button without data.</summary>
    public SlicerStyleElementOptions? HoveredSelectedItemWithNoData { get; set; }
    /// <summary>Hovered unselected button with data.</summary>
    public SlicerStyleElementOptions? HoveredUnselectedItemWithData { get; set; }
    /// <summary>Hovered unselected button without data.</summary>
    public SlicerStyleElementOptions? HoveredUnselectedItemWithNoData { get; set; }

    internal TableStyleOptions ToTableStyleOptions()
    {
        var result = new TableStyleOptions { ShowAsAvailableSlicerStyle = true, Elements = [] };
        Add(WholeTable, "xlWholeTable");
        Add(Header, "xlHeaderRow");
        Add(SelectedItemWithData, "xlSlicerSelectedItemWithData");
        Add(SelectedItemWithNoData, "xlSlicerSelectedItemWithNoData");
        Add(UnselectedItemWithData, "xlSlicerUnselectedItemWithData");
        Add(UnselectedItemWithNoData, "xlSlicerUnselectedItemWithNoData");
        Add(HoveredSelectedItemWithData, "xlSlicerHoveredSelectedItemWithData");
        Add(HoveredSelectedItemWithNoData, "xlSlicerHoveredSelectedItemWithNoData");
        Add(HoveredUnselectedItemWithData, "xlSlicerHoveredUnselectedItemWithData");
        Add(HoveredUnselectedItemWithNoData, "xlSlicerHoveredUnselectedItemWithNoData");
        return result;

        void Add(SlicerStyleElementOptions? format, string element)
        {
            if (format is null) return;
            result.Elements.Add(new TableStyleElementOptions
            {
                ElementType = element,
                FillColor = format.FillColor,
                FontColor = format.FontColor,
                Bold = format.Bold,
                Borders = format.BorderColor is null ? null :
                    new[] { CellBorderPosition.Left, CellBorderPosition.Top, CellBorderPosition.Bottom, CellBorderPosition.Right }
                        .Select(position => new CellBorderOptions
                        {
                            Position = position,
                            Color = format.BorderColor,
                            LineStyle = "continuous",
                            Weight = "thin"
                        }).ToList()
            });
        }
    }
}

/// <summary>Selected overrides for one slicer-style element; omitted fields retain their source values.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class SlicerStyleElementOptions
{
    /// <summary>Solid fill, #RRGGBB.</summary>
    public string? FillColor { get; set; }
    /// <summary>Four solid, thin outer borders, #RRGGBB.</summary>
    public string? BorderColor { get; set; }
    /// <summary>Font colour, #RRGGBB.</summary>
    public string? FontColor { get; set; }
    /// <summary>Bold state; false explicitly removes bold.</summary>
    public bool? Bold { get; set; }
}

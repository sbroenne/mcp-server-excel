using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Changes only supplied settings on an existing rule. Its native type is retained.
/// Null means leave unchanged; visual settings must match the selected rule type.
/// </summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class ConditionalRuleUpdateOptions
{
    /// <summary>Exact worksheet range, named range, or disjoint scope.</summary>
    public string? AppliesTo { get; set; }
    /// <summary>Stop evaluating lower-priority rules when true; unavailable for visual scales/bars/icons.</summary>
    public bool? StopIfTrue { get; set; }
    /// <summary>Comparison operator for a cell-value rule.</summary>
    public string? OperatorType { get; set; }
    /// <summary>First formula/value for a cell-value or expression rule.</summary>
    public string? Formula1 { get; set; }
    /// <summary>Second formula/value for between/notBetween cell-value rules.</summary>
    public string? Formula2 { get; set; }
    /// <summary>Fill color as #RRGGBB or color index.</summary>
    public string? InteriorColor { get; set; }
    /// <summary>Native fill pattern.</summary>
    public string? InteriorPattern { get; set; }
    /// <summary>Font color as #RRGGBB or color index.</summary>
    public string? FontColor { get; set; }
    /// <summary>Whether the rule applies bold text.</summary>
    public bool? FontBold { get; set; }
    /// <summary>Whether the rule applies italic text.</summary>
    public bool? FontItalic { get; set; }
    /// <summary>Border style for all four conditional-format borders.</summary>
    public string? BorderStyle { get; set; }
    /// <summary>Border color as #RRGGBB or color index.</summary>
    public string? BorderColor { get; set; }
    /// <summary>Color-scale minimum threshold type.</summary>
    public string? ColorScaleMinType { get; set; }
    /// <summary>Color-scale minimum threshold value.</summary>
    public string? ColorScaleMinValue { get; set; }
    /// <summary>Color-scale minimum stop color.</summary>
    public string? ColorScaleMinColor { get; set; }
    /// <summary>Color-scale midpoint threshold type; existing three-stop scales only.</summary>
    public string? ColorScaleMidType { get; set; }
    /// <summary>Color-scale midpoint threshold value.</summary>
    public string? ColorScaleMidValue { get; set; }
    /// <summary>Color-scale midpoint stop color.</summary>
    public string? ColorScaleMidColor { get; set; }
    /// <summary>Color-scale maximum threshold type.</summary>
    public string? ColorScaleMaxType { get; set; }
    /// <summary>Color-scale maximum threshold value.</summary>
    public string? ColorScaleMaxValue { get; set; }
    /// <summary>Color-scale maximum stop color.</summary>
    public string? ColorScaleMaxColor { get; set; }
    /// <summary>Positive data-bar fill color.</summary>
    public string? DataBarColor { get; set; }
    /// <summary>Custom negative data-bar fill color.</summary>
    public string? DataBarNegativeColor { get; set; }
    /// <summary>Data-bar direction: context, leftToRight, rightToLeft.</summary>
    public string? DataBarDirection { get; set; }
    /// <summary>Whether to display the value beside the data bar.</summary>
    public bool? DataBarShowValue { get; set; }
    /// <summary>Data-bar minimum threshold type.</summary>
    public string? DataBarMinType { get; set; }
    /// <summary>Data-bar minimum threshold value.</summary>
    public string? DataBarMinValue { get; set; }
    /// <summary>Data-bar maximum threshold type.</summary>
    public string? DataBarMaxType { get; set; }
    /// <summary>Data-bar maximum threshold value.</summary>
    public string? DataBarMaxValue { get; set; }
    /// <summary>Native icon-set name; changing sets may reset native thresholds.</summary>
    public string? IconSetId { get; set; }
    /// <summary>Whether to reverse icon order.</summary>
    public bool? IconSetReverse { get; set; }
    /// <summary>Whether to hide the value and display only its icon.</summary>
    public bool? IconSetShowIconOnly { get; set; }
    /// <summary>First editable icon threshold type; the implicit lower bound is retained.</summary>
    public string? IconThreshold1Type { get; set; }
    /// <summary>First editable icon threshold value.</summary>
    public string? IconThreshold1Value { get; set; }
    /// <summary>Second editable icon threshold type.</summary>
    public string? IconThreshold2Type { get; set; }
    /// <summary>Second editable icon threshold value.</summary>
    public string? IconThreshold2Value { get; set; }
    /// <summary>Third editable icon threshold type; four/five-icon sets only.</summary>
    public string? IconThreshold3Type { get; set; }
    /// <summary>Third editable icon threshold value.</summary>
    public string? IconThreshold3Value { get; set; }
    /// <summary>Fourth editable icon threshold type; five-icon sets only.</summary>
    public string? IconThreshold4Type { get; set; }
    /// <summary>Fourth editable icon threshold value.</summary>
    public string? IconThreshold4Value { get; set; }
    /// <summary>Top/bottom rank: 1-1000 values or 1-100 percent.</summary>
    public int? Rank { get; set; }
    /// <summary>Whether the top/bottom rank is a percentage.</summary>
    public bool? Top10Percent { get; set; }
    /// <summary>Top/bottom direction: top or bottom.</summary>
    public string? TopBottom { get; set; }
    /// <summary>Above/below-average comparison selector.</summary>
    public string? AboveBelow { get; set; }
    /// <summary>Date-period selector for an existing time-period rule.</summary>
    public string? DatePeriod { get; set; }
    /// <summary>Whether an existing unique-values rule highlights duplicates.</summary>
    public bool? DuplicateValues { get; set; }
    /// <summary>Standard-deviation multiplier: 1-3 for an above/below-standard-deviation rule.</summary>
    public int? StandardDeviations { get; set; }
}

#pragma warning disable CS1591
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Core.Models.Actions;

/// <summary>
/// Actions available for file tool
/// </summary>
/// <remarks>
/// IMPORTANT: Keep enum values synchronized with tool switch cases.
/// Enum names are PascalCase (e.g., Create), serialized as kebab-case (e.g., create) via JsonStringEnumMemberName.
/// ActionExtensions.ToActionString() also returns kebab-case for logging/routing.
/// </remarks>
[System.Text.Json.Serialization.JsonConverter(typeof(System.Text.Json.Serialization.JsonStringEnumConverter<FileAction>))]
public enum FileAction
{
    /// <summary>Lists sessions owned by the current entry-point process.</summary>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [System.Text.Json.Serialization.JsonStringEnumMemberName("list")]
    List,

    /// <summary>Opens one exact existing workbook through the platform session owner.</summary>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [System.Text.Json.Serialization.JsonStringEnumMemberName("open")]
    Open,

    /// <summary>Closes one owned workbook with explicit save or discard semantics.</summary>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [System.Text.Json.Serialization.JsonStringEnumMemberName("close")]
    Close,

    /// <summary>Creates and owns a workbook at one exact target path.</summary>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [System.Text.Json.Serialization.JsonStringEnumMemberName("create")]
    Create,

    /// <summary>Checks whether a workbook path can be used without opening it.</summary>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [System.Text.Json.Serialization.JsonStringEnumMemberName("test")]
    Test
}

// NOTE: PowerQueryAction is now generated from IPowerQueryCommands interface
// See Sbroenne.ExcelMcp.Generated.PowerQueryAction in ServiceRegistry.PowerQuery.g.cs

// NOTE: SheetAction and SheetStyleAction are now generated from ISheetCommands and ISheetStyleCommands
// See Sbroenne.ExcelMcp.Generated.SheetAction in ServiceRegistry.Sheet.g.cs
// See Sbroenne.ExcelMcp.Generated.SheetStyleAction in ServiceRegistry.SheetStyle.g.cs

// NOTE: RangeAction is now generated from IRangeCommands interface
// See Sbroenne.ExcelMcp.Generated.RangeAction in ServiceRegistry.Range.g.cs

// NOTE: RangeEditAction is now generated from IRangeEditCommands interface
// See Sbroenne.ExcelMcp.Generated.RangeEditAction in ServiceRegistry.RangeEdit.g.cs

// NOTE: RangeFormatAction is now generated from IRangeFormatCommands interface
// See Sbroenne.ExcelMcp.Generated.RangeFormatAction in ServiceRegistry.RangeFormat.g.cs

// NOTE: RangeLinkAction is now generated from IRangeLinkCommands interface
// See Sbroenne.ExcelMcp.Generated.RangeLinkAction in ServiceRegistry.RangeLink.g.cs

// NamedRangeAction is now generated in ServiceRegistry.NamedRange

// ConditionalFormatAction is now generated in ServiceRegistry.ConditionalFormat

// VbaAction is now generated from IVbaCommands interface
// See Sbroenne.ExcelMcp.Generated.VbaAction in ServiceRegistry.Vba.g.cs

// ConnectionAction is now generated in ServiceRegistry.Connection

// DataModelAction is now generated in ServiceRegistry.DataModel
// DataModelRelAction is now generated in ServiceRegistry.DataModelRel

// TableAction is now generated in ServiceRegistry.Table
// TableColumnAction is now generated in ServiceRegistry.TableColumn

// PivotTableAction is now generated in ServiceRegistry.PivotTable
// PivotTableFieldAction is now generated in ServiceRegistry.PivotTableField
// PivotTableCalcAction is now generated in ServiceRegistry.PivotTableCalc

// ChartAction is now generated from IChartCommands interface
// See Sbroenne.ExcelMcp.Generated.ChartAction in ServiceRegistry.Chart.g.cs

// ChartConfigAction is now generated from IChartConfigCommands interface
// See Sbroenne.ExcelMcp.Generated.ChartConfigAction in ServiceRegistry.ChartConfig.g.cs

// SlicerAction is now generated from ISlicerCommands interface
// See Sbroenne.ExcelMcp.Generated.SlicerAction in ServiceRegistry.Slicer.g.cs

// CalculationModeAction is now generated from ICalculationModeCommands interface
// See Sbroenne.ExcelMcp.Generated.CalculationAction in ServiceRegistry.Calculation.g.cs
#pragma warning restore CS1591

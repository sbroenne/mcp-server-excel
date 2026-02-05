// Suppress IDE0005 (unnecessary using) – explicit usings kept for clarity in test reflection code
#pragma warning disable IDE0005
using System.Reflection;
#pragma warning restore IDE0005
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// CRITICAL: Automated verification that all Core Commands methods are exposed via MCP actions.
/// These tests PREVENT regression by ensuring compile-time and runtime coverage.
///
/// If these tests fail, it means:
/// 1. A new Core method was added but no MCP action was created
/// 2. An enum value is missing from ToolActions.cs
/// 3. A ToActionString mapping is missing from ActionExtensions.cs
///
/// DO NOT disable or skip these tests without fixing the underlying coverage gap!
/// </summary>
public class CoreCommandsCoverageTests
{
    /// <summary>
    /// Verifies IPowerQueryCommands has matching PowerQueryAction enum values
    /// </summary>
    [Fact]
    public void IPowerQueryCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IPowerQueryCommands));
        var enumValueCount = Enum.GetValues<PowerQueryAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IPowerQueryCommands has {coreMethodCount} methods but PowerQueryAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to ToolActions.cs!");
    }

    /// <summary>
    /// Verifies ISheetCommands has matching SheetAction enum values
    /// </summary>
    [Fact]
    public void ISheetCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(ISheetCommands));
        var enumValueCount = Enum.GetValues<SheetAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"ISheetCommands has {coreMethodCount} methods but SheetAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to interface or regenerate!");
    }

    /// <summary>
    /// Verifies IRangeCommands has matching RangeAction enum values
    /// </summary>
    [Fact]
    public void IRangeCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IRangeCommands));
        var enumValueCount = Enum.GetValues<RangeAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IRangeCommands has {coreMethodCount} methods but RangeAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to ToolActions.cs!");
    }

    /// <summary>
    /// Verifies ITableCommands has matching TableAction enum values
    /// </summary>
    [Fact]
    public void ITableCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(ITableCommands));
        var enumValueCount = Enum.GetValues<TableAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"ITableCommands has {coreMethodCount} methods but TableAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to ToolActions.cs!");
    }

    /// <summary>
    /// Verifies IConnectionCommands has matching ConnectionAction enum values
    /// </summary>
    [Fact]
    public void IConnectionCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IConnectionCommands));
        var enumValueCount = Enum.GetValues<ConnectionAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IConnectionCommands has {coreMethodCount} methods but ConnectionAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to ToolActions.cs!");
    }

    /// <summary>
    /// Verifies IDataModelCommands has matching DataModelAction enum values
    /// </summary>
    [Fact]
    public void IDataModelCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IDataModelCommands));
        var enumValueCount = Enum.GetValues<DataModelAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IDataModelCommands has {coreMethodCount} methods but DataModelAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to ToolActions.cs!");
    }

    /// <summary>
    /// Verifies IPivotTableCommands has matching PivotTableAction enum values
    /// </summary>
    [Fact]
    public void IPivotTableCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IPivotTableCommands));
        var enumValueCount = Enum.GetValues<PivotTableAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IPivotTableCommands has {coreMethodCount} methods but PivotTableAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to interface or regenerate!");
    }

    /// <summary>
    /// Verifies IPivotTableFieldCommands has matching PivotTableFieldAction enum values
    /// </summary>
    [Fact]
    public void IPivotTableFieldCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IPivotTableFieldCommands));
        var enumValueCount = Enum.GetValues<PivotTableFieldAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IPivotTableFieldCommands has {coreMethodCount} methods but PivotTableFieldAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to interface or regenerate!");
    }

    /// <summary>
    /// Verifies IPivotTableCalcCommands has matching PivotTableCalcAction enum values
    /// </summary>
    [Fact]
    public void IPivotTableCalcCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IPivotTableCalcCommands));
        var enumValueCount = Enum.GetValues<PivotTableCalcAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IPivotTableCalcCommands has {coreMethodCount} methods but PivotTableCalcAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to interface or regenerate!");
    }

    /// <summary>
    /// Verifies ISlicerCommands has matching SlicerAction enum values
    /// </summary>
    [Fact]
    public void ISlicerCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(ISlicerCommands));
        var enumValueCount = Enum.GetValues<SlicerAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"ISlicerCommands has {coreMethodCount} methods but SlicerAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to interface or regenerate!");
    }

    /// <summary>
    /// Verifies IChartCommands has matching ChartAction enum values
    /// </summary>
    [Fact]
    public void IChartCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IChartCommands));
        var enumValueCount = Enum.GetValues<ChartAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IChartCommands has {coreMethodCount} methods but ChartAction has only {enumValueCount} enum values. " +
            $"Add missing enum values or regenerate!");
    }

    /// <summary>
    /// Verifies IChartConfigCommands has matching ChartConfigAction enum values
    /// </summary>
    [Fact]
    public void IChartConfigCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(IChartConfigCommands));
        var enumValueCount = Enum.GetValues<ChartConfigAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"IChartConfigCommands has {coreMethodCount} methods but ChartConfigAction has only {enumValueCount} enum values. " +
            $"Add missing enum values or regenerate!");
    }

    /// <summary>
    /// Verifies all PowerQueryAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void PowerQueryAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<PowerQueryAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.PowerQuery.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.PowerQuery.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all SheetAction enum values have ToActionString mappings (via generated ServiceRegistry)
    /// </summary>
    [Fact]
    public void SheetAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SheetAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Sheet.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Sheet.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all SheetStyleAction enum values have ToActionString mappings (via generated ServiceRegistry)
    /// </summary>
    [Fact]
    public void SheetStyleAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SheetStyleAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.SheetStyle.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.SheetStyle.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all RangeAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void RangeAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<RangeAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Range.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Range.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all TableAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void TableAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<TableAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Table.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Table.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all ConnectionAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void ConnectionAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ConnectionAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Connection.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Connection.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all DataModelAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void DataModelAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<DataModelAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.DataModel.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.DataModel.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all PivotTableAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void PivotTableAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<PivotTableAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.PivotTable.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.PivotTable.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all PivotTableFieldAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void PivotTableFieldAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<PivotTableFieldAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.PivotTableField.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.PivotTableField.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all PivotTableCalcAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void PivotTableCalcAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<PivotTableCalcAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.PivotTableCalc.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.PivotTableCalc.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all ChartAction enum values have ToActionString mappings (via generated ServiceRegistry)
    /// </summary>
    [Fact]
    public void ChartAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ChartAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Chart.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Chart.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all ChartConfigAction enum values have ToActionString mappings (via generated ServiceRegistry)
    /// </summary>
    [Fact]
    public void ChartConfigAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ChartConfigAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.ChartConfig.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.ChartConfig.ToActionString(action));
        }
    }

    /// <summary>
    /// Verifies all SlicerAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void SlicerAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SlicerAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Slicer.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Slicer.ToActionString(action));
        }
    }

    /// <summary>
    /// Helper: Counts public async methods in an interface (excludes properties, events, etc.)
    /// </summary>
    private static int GetAsyncMethodCount(Type interfaceType)
    {
        // Count DISTINCT async method base names (treat overloads as single logical operation).
        // Reason: Enum actions represent semantic operations, not overload variants.
        // Example: Refresh(...) and Refresh(..., TimeSpan?) map to single "refresh" action.
        return interfaceType
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => m.Name.EndsWith("Async", StringComparison.Ordinal))
            .Select(m => m.Name) // includes overload name twice
            .Distinct(StringComparer.Ordinal) // collapse overloads
            .Count();
    }
}





// Suppress IDE0005 (unnecessary using) – explicit usings kept for clarity in test reflection code
#pragma warning disable IDE0005
using System.Reflection;
#pragma warning restore IDE0005
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.McpServer.Models;
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
    /// Verifies ISheetCommands has matching WorksheetAction enum values
    /// </summary>
    [Fact]
    public void ISheetCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetAsyncMethodCount(typeof(ISheetCommands));
        var enumValueCount = Enum.GetValues<WorksheetAction>().Length;

        Assert.True(
            enumValueCount >= coreMethodCount,
            $"ISheetCommands has {coreMethodCount} methods but WorksheetAction has only {enumValueCount} enum values. " +
            $"Add missing enum values to ToolActions.cs!");
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
            $"Add missing enum values to ToolActions.cs!");
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
            $"Add missing enum values to ToolActions.cs!");
    }

    /// <summary>
    /// Verifies all PowerQueryAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void PowerQueryAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<PowerQueryAction>())
        {
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
        }
    }

    /// <summary>
    /// Verifies all WorksheetAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void WorksheetAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<WorksheetAction>())
        {
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
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
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
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
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
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
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
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
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
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
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
        }
    }

    /// <summary>
    /// Verifies all ChartAction enum values have ToActionString mappings
    /// </summary>
    [Fact]
    public void ChartAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ChartAction>())
        {
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
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
            var exception = Record.Exception(() => action.ToActionString());
            Assert.Null(exception);
            Assert.NotEmpty(action.ToActionString());
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

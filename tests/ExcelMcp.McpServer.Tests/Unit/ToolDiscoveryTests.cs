using System.Reflection;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

/// <summary>
/// Tests to verify that all MCP tools are properly decorated with required attributes
/// for discovery by the MCP SDK's WithToolsFromAssembly() method.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ToolDiscovery")]
public class ToolDiscoveryTests
{
    [Fact]
    public void ExcelPivotTableTool_HasMcpServerToolTypeAttribute()
    {
        // Arrange
        var toolType = typeof(ExcelPivotTableTool);

        // Act
        var attribute = toolType.GetCustomAttribute<McpServerToolTypeAttribute>();

        // Assert
        Assert.NotNull(attribute);
    }

    [Fact]
    public void ExcelPivotTableTool_HasMcpServerToolAttributeWithName()
    {
        // Arrange
        var toolType = typeof(ExcelPivotTableTool);
        var method = toolType.GetMethod("ExcelPivotTable", BindingFlags.Public | BindingFlags.Static);
        Assert.NotNull(method);

        // Act
        var attribute = method!.GetCustomAttribute<McpServerToolAttribute>();

        // Assert
        Assert.NotNull(attribute);
        Assert.Equal("excel_pivottable", attribute!.Name);
    }

    [Theory]
    [InlineData(typeof(ExcelConnectionTool), "ExcelConnection", "excel_connection")]
    [InlineData(typeof(ExcelDataModelTool), "ExcelDataModel", "excel_datamodel")]
    [InlineData(typeof(ExcelFileTool), "ExcelFile", "excel_file")]
    [InlineData(typeof(ExcelNamedRangeTool), "ExcelParameter", "excel_namedrange")]
    [InlineData(typeof(ExcelPivotTableTool), "ExcelPivotTable", "excel_pivottable")]
    [InlineData(typeof(ExcelPowerQueryTool), "ExcelPowerQuery", "excel_powerquery")]
    [InlineData(typeof(ExcelQueryTableTool), "ExcelQueryTable", "excel_querytable")]
    [InlineData(typeof(ExcelRangeTool), "ExcelRange", "excel_range")]
    [InlineData(typeof(TableTool), "Table", "excel_table")]
    [InlineData(typeof(ExcelVbaTool), "ExcelVba", "excel_vba")]
    [InlineData(typeof(ExcelWorksheetTool), "ExcelWorksheet", "excel_worksheet")]
    public void AllTools_HaveMcpServerToolAttributeWithCorrectName(Type toolType, string methodName, string expectedToolName)
    {
        // Arrange
        var method = toolType.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static);
        Assert.NotNull(method);

        // Act
        var attribute = method!.GetCustomAttribute<McpServerToolAttribute>();

        // Assert
        Assert.NotNull(attribute);
        Assert.Equal(expectedToolName, attribute!.Name);
    }

    [Theory]
    [InlineData(typeof(ExcelConnectionTool))]
    [InlineData(typeof(ExcelDataModelTool))]
    [InlineData(typeof(ExcelFileTool))]
    [InlineData(typeof(ExcelNamedRangeTool))]
    [InlineData(typeof(ExcelPivotTableTool))]
    [InlineData(typeof(ExcelPowerQueryTool))]
    [InlineData(typeof(ExcelQueryTableTool))]
    [InlineData(typeof(ExcelRangeTool))]
    [InlineData(typeof(TableTool))]
    [InlineData(typeof(ExcelVbaTool))]
    [InlineData(typeof(ExcelWorksheetTool))]
    public void AllTools_HaveMcpServerToolTypeAttribute(Type toolType)
    {
        // Act
        var attribute = toolType.GetCustomAttribute<McpServerToolTypeAttribute>();

        // Assert
        Assert.NotNull(attribute);
    }

    /// <summary>
    /// Tests that all expected tools are discoverable via assembly scanning,
    /// simulating what the MCP SDK's WithToolsFromAssembly() does.
    /// This catches issues like partial classes that prevent runtime discovery.
    /// </summary>
    [Fact]
    public void AssemblyScan_DiscoversAllExpectedTools()
    {
        // Arrange - Expected tool names that should be discoverable
        var expectedToolNames = new HashSet<string>
        {
            "excel_conditionalformat",
            "excel_connection",
            "excel_datamodel",
            "excel_file",
            "excel_namedrange",
            "excel_pivottable",
            "excel_powerquery",
            "excel_querytable",
            "excel_range",
            "excel_table",
            "excel_vba",
            "excel_worksheet"
        };
        // Act - Scan assembly for tool types (simulating MCP SDK behavior)
        var assembly = typeof(ExcelPivotTableTool).Assembly;
        var toolTypes = assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .ToList();

        // Extract tool names from discovered types
        var discoveredToolNames = new HashSet<string>();
        foreach (var toolType in toolTypes)
        {
            var methods = toolType.GetMethods(BindingFlags.Public | BindingFlags.Static);
            foreach (var method in methods)
            {
                var toolAttr = method.GetCustomAttribute<McpServerToolAttribute>();
                if (toolAttr?.Name != null)
                {
                    discoveredToolNames.Add(toolAttr.Name);
                }
            }
        }

        // Assert - All expected tools must be discovered
        Assert.Equal(expectedToolNames.Count, discoveredToolNames.Count);

        var missingTools = expectedToolNames.Except(discoveredToolNames).ToList();
        Assert.Empty(missingTools); // This would have failed before the fix!

        var unexpectedTools = discoveredToolNames.Except(expectedToolNames).ToList();
        Assert.Empty(unexpectedTools);
    }
}

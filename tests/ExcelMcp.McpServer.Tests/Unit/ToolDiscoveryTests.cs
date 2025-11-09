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
}

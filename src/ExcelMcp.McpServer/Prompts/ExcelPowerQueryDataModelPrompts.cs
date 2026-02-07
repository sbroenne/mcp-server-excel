using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for teaching LLMs about Power Query, Data Model, and Table workflows.
/// Critical upfront knowledge to prevent common mistakes.
/// </summary>
[McpServerPromptType]
public static class ExcelPowerQueryDataModelPrompts
{
    [McpServerPrompt(Name = "excel_powerquery_guide")]
    [Description("Power Query M code, load destinations, and data import workflows")]
    public static ChatMessage PowerQueryGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_powerquery.md"));
    }

    [McpServerPrompt(Name = "excel_table_guide")]
    [Description("Excel Tables: lifecycle, data operations, and adding tables to Power Pivot Data Model")]
    public static ChatMessage TableGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_table.md"));
    }

    [McpServerPrompt(Name = "excel_datamodel_guide")]
    [Description("Data Model (Power Pivot): DAX measures, table management, and prerequisites for analysis")]
    public static ChatMessage DataModelGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_datamodel.md"));
    }

    [McpServerPrompt(Name = "excel_chart_guide")]
    [Description("Chart operations: PivotCharts vs regular charts, charting Data Model data")]
    public static ChatMessage ChartGuide()
    {
        return new ChatMessage(ChatRole.User, MarkdownLoader.LoadPrompt("excel_chart.md"));
    }
}



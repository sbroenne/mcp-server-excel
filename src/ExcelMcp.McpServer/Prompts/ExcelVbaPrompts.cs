using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel VBA macro management.
/// </summary>
[McpServerPromptType]
public static class ExcelVbaPrompts
{
    /// <summary>
    /// Guide for VBA macro version control and automation workflows.
    /// </summary>
    [McpServerPrompt(Name = "excel_vba_version_control_guide")]
    [Description("Guide for managing VBA macros with version control and automation")]
    public static ChatMessage VbaVersionControlGuide()
    {
        return new ChatMessage(ChatRole.User, @"When working with Excel VBA macros, use excel_vba tool for version control and automation.

# VBA VERSION CONTROL WORKFLOW

## EXPORT MACROS FOR GIT

User wants to: ""Save VBA code to version control""

1. List all modules: excel_vba(action: 'list', excelPath: 'workbook.xlsm')
2. Export each module: excel_vba(action: 'export', excelPath: 'workbook.xlsm', moduleName: 'Module1', targetPath: 'vba/Module1.bas')

TIP: Use batch mode for multiple modules!

## IMPORT MACROS FROM GIT

User wants to: ""Load VBA code from repository""

1. Import module: excel_vba(action: 'import', excelPath: 'workbook.xlsm', moduleName: 'Module1', sourcePath: 'vba/Module1.bas')
2. Verify: excel_vba(action: 'view', excelPath: 'workbook.xlsm', moduleName: 'Module1')

## RUN MACROS FOR AUTOMATION

User wants to: ""Execute VBA macro""

excel_vba(action: 'run', excelPath: 'workbook.xlsm', moduleName: 'Module1', macroName: 'ProcessData')

# COMMON PATTERNS

**CI/CD Integration:**
- Export macros before commit
- Import macros after checkout
- Run validation macros in pipeline

**Team Collaboration:**
- Export: Share VBA code via Git
- Import: Sync VBA from teammates
- Update: Modify VBA programmatically

**Macro Automation:**
- Run: Execute data processing macros
- List: Discover available macros
- View: Review macro code

# IMPORTANT NOTES

- **File format**: Use .xlsm for macro-enabled workbooks
- **Security**: Trust settings may block macro execution
- **VBA Editor**: For complex editing, use Excel UI
- **Parameters**: Pass parameters to macros as comma-separated values");
    }
}

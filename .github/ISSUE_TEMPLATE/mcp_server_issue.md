---
name: MCP Server Issue
about: Report issues with the MCP Server for AI assistants
title: '[MCP] '
labels: 'mcp-server'
assignees: ''

---

## Issue Description
A clear and concise description of the MCP Server issue.

## AI Assistant
Which AI assistant are you using with the MCP Server?
- [ ] **GitHub Copilot** (VS Code, Visual Studio, etc.)
- [ ] **Claude Desktop** (Anthropic)
- [ ] **ChatGPT** (OpenAI)
- [ ] **Other**: [please specify]

## MCP Tool & Action
Which MCP tool and action are experiencing issues?
- **Tool**: [e.g., powerquery, worksheet, vba, excel_cell, excel_parameter, file]
- **Action**: [e.g., list, view, import, export, update, refresh, delete, etc.]
- **File Path**: [e.g., "C:\Data\workbook.xlsx"]
- **Additional Parameters**: [describe any other parameters used]

## Expected Behavior
What did you expect the MCP Server to do?

## Actual Behavior
What did the MCP Server actually do?

## Error Response
If you received an error, paste the full JSON response:
```json
{
  "error": "paste error here"
}
```

## MCP Server Configuration
How is the MCP Server configured?

**Configuration file location**: [e.g., `.config/Code/User/globalStorage/github.copilot-chat/config.json`]

**MCP Configuration**:
```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "mcp-excel",
      // or other configuration
    }
  }
}
```

## Environment
- **Windows Version**: [e.g. Windows 11, Windows 10]
- **Excel Version**: [e.g. Excel 365, Excel 2019]
- **ExcelMcp Version**: [e.g. v1.0.0 - run `mcp-excel --version` or `dotnet tool list -g`]
- **.NET Version**: [Run `dotnet --version`]
- **Installation Method**: 
  - [ ] Global .NET tool (`dotnet tool install --global ExcelMcp.McpServer`)
  - [ ] Source build
  - [ ] Other: [please specify]

## MCP Server Logs
If possible, provide relevant logs from the MCP Server:
```
[Paste logs here]
```

## Steps to Reproduce
1. Configure AI assistant with MCP Server
2. Ask AI assistant: "..."
3. MCP Server receives request for tool: [tool_name], action: [action_name]
4. See error

## Conversation Context (Optional)
If helpful, provide the conversation you had with the AI assistant that led to this issue:
```
User: "Can you list all Power Queries in my workbook?"
AI: [response]
[MCP Server error occurs]
```

## Excel File Details
- **File Format**: [.xlsx or .xlsm]
- **File Size**: [approximate size]
- **Contains**: 
  - [ ] Power Queries
  - [ ] VBA Macros
  - [ ] Named Ranges
  - [ ] Multiple worksheets
  - [ ] External connections

## VBA-Related Issues (if applicable)
- [ ] VBA trust is properly configured (`ExcelMcp check-vba-trust`)
- [ ] Using .xlsm file format for VBA operations
- [ ] VBA module exists in the workbook
- [ ] Macro security settings allow programmatic access

## Additional Context
Add any other context about the problem here, including:
- Screenshots of AI assistant interaction
- Sample Excel files (with sensitive data removed)
- Other relevant information

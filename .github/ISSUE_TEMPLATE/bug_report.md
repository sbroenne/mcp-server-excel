---
name: Bug Report
about: Create a report to help us improve ExcelMcp
title: '[BUG] '
labels: 'bug'
assignees: ''

---

## Bug Description
A clear and concise description of what the bug is.

## Component
Which component is this bug related to?
- [ ] **MCP Server** (Model Context Protocol server for AI assistants - `mcp-excel`)
- [ ] **CLI** (Command-line interface - `ExcelMcp.exe`)
- [ ] **Core Library** (Shared functionality)
- [ ] **Not sure**

## Command/Usage
**For CLI:**
```
ExcelMcp <command> <arguments>
```

**For MCP Server:**
- Tool name: [e.g., powerquery, worksheet, etc.]
- Action: [e.g., list, view, import, etc.]
- Parameters used: [describe what was passed]

## Expected Behavior
A clear and concise description of what you expected to happen.

## Actual Behavior
A clear and concise description of what actually happened.

## Error Message
If applicable, paste the full error message:
```
[Error message here]
```

## Environment
- **Windows Version**: [e.g. Windows 11, Windows 10]
- **Excel Version**: [e.g. Excel 365, Excel 2019]
- **ExcelMcp Version**: [e.g. v1.0.0]
- **.NET Version**: [Run `dotnet --version`]
- **Installation Method**: [NuGet tool / Binary download / Source build]
- **File Format**: [e.g. .xlsx, .xlsm]
- **VBA Trust Enabled**: [Yes/No - if VBA-related issue]
- **AI Assistant** (if using MCP Server): [e.g., GitHub Copilot, Claude Desktop, ChatGPT, etc.]

## Sample File
If possible, attach a sample Excel file that reproduces the issue (remove sensitive data).

## VBA-Related Issues (if applicable)
- [ ] VBA trust is properly configured (`ExcelMcp check-vba-trust`)
- [ ] Using .xlsm file format for VBA commands
- [ ] VBA module exists in the workbook
- [ ] Macro security settings allow programmatic access

## Steps to Reproduce
1. Go to '...'  
2. Click on '....'
3. Scroll down to '....'
4. See error

## Additional Context
Add any other context about the problem here.

## Excel Process Cleanup
- [ ] Excel processes clean up properly after the command
- [ ] Excel processes remain running (this is part of the bug)
- [ ] Not applicable/unsure
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
- [ ] **CLI** (Command-line interface - `excelcli`)
- [ ] **Core Library** (Shared functionality)
- [ ] **Not sure**

## Command/Usage
**For CLI:**
```text
excelcli <command> <action> <arguments>
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
```text
[Error message here]
```

## Environment
- **OS and Architecture**: [e.g. Windows 11 x64, macOS on Apple Silicon]
- **Excel Version**: [e.g. Excel 365, Excel 2019]
- **ExcelMcp Version**: [e.g. v1.0.0]
- **.NET Version**: [Only for .NET tool/source builds; otherwise N/A]
- **Node.js Version**: [For npm/plugin installations; otherwise N/A]
- **Installation Method**: [npm / VS Code extension / MCPB / Copilot plugin / .NET tool / Binary download / Source build]
- **File Format**: [e.g. .xlsx, .xlsm]
- **VBA Trust Enabled**: [Yes/No - Windows VBA issues only; VBA is unsupported in the Mac beta]
- **AI Assistant** (if using MCP Server): [e.g., GitHub Copilot, Claude Desktop, ChatGPT, etc.]

## Sample File
If possible, attach a sample Excel file that reproduces the issue (remove sensitive data).

## VBA-Related Issues (if applicable)
- [ ] Excel Trust Center setting "Trust access to the VBA project object model" is enabled
- [ ] Using .xlsm file format for VBA commands
- [ ] VBA module exists in the workbook
- [ ] Macro security settings allow programmatic access

## Steps to Reproduce
List the exact commands, tool calls, workbook setup, and steps that reproduce
the problem.

## Additional Context
Add any other context about the problem here.

## Excel Process Cleanup
- [ ] Windows: the session-owned Excel process cleans up properly
- [ ] Windows: the session-owned Excel process remains running
- [ ] Mac: the confirmed session-owned workbook closes; shared Excel and unrelated workbooks remain untouched
- [ ] Mac: an uncertain open reports RecoveryRequired and retains the exact workbook for manual reconciliation
- [ ] Not applicable/unsure

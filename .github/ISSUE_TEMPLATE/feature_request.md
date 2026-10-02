---
name: Feature Request
about: Suggest an idea for ExcelMcp
title: '[FEATURE] '
labels: 'enhancement'
assignees: ''

---

## Is your feature request related to a problem?
A clear and concise description of what the problem is. Ex. I'm always frustrated when [...]

## Component
Which component should this feature be added to?
- [ ] **MCP Server** (AI assistant integration)
- [ ] **CLI** (Command-line interface)
- [ ] **Both** (MCP Server and CLI)
- [ ] **Core Library** (Shared functionality)
- [ ] **Not sure**

## Describe the solution you'd like
A clear and concise description of what you want to happen.

## Proposed Syntax

**For CLI:**
```bash
excelcli new-command <parameters>
```

**For MCP Server:**
- Tool: [e.g., powerquery, worksheet, vba]
- Action: [e.g., new-action]
- Parameters: [describe expected parameters]

## Describe alternatives you've considered
A clear and concise description of any alternative solutions or features you've considered.

## Use Case
Describe the specific use case this feature would address:
- [ ] Power Query management
- [ ] VBA script automation
- [ ] Worksheet operations  
- [ ] Data analysis
- [ ] Coding agent automation
- [ ] Macro-enabled workbook (.xlsm) operations
- [ ] Other: [please specify]

## Target Users
Who would benefit from this feature?
- [ ] **AI Assistants** (GitHub Copilot, Claude, ChatGPT via MCP Server)
- [ ] **Direct CLI Users** (Command-line automation)
- [ ] **Desktop CI/CD Pipelines** (Interactive Excel hosts; headless hosts are unsupported)
- [ ] **Excel Developers** (Power Query, VBA development)
- [ ] **Data Engineers** (ETL workflows)
- [ ] Other: [please specify]

## Excel Operations Involved
Specify whether the request targets Windows, the experimental Apple Silicon
Mac beta, or both. Existing COM APIs do not imply Mac support.

What Excel APIs or operations would this feature likely use?
- [ ] Power Query (Workbook.Queries)
- [ ] Worksheets (Worksheet operations)
- [ ] Named Ranges (Workbook.Names)
- [ ] Charts/PivotTables
- [ ] VBA/Macros (VBProject.VBComponents)
- [ ] External connections
- [ ] Macro-enabled workbooks (.xlsm)
- [ ] Other: [please specify]

## Additional context
Add any other context, screenshots, or examples about the feature request here.

## Implementation Notes
If you have ideas about how this could be implemented, please share them here.
---
name: excel-cli
description: >
  Automate Microsoft Excel on Windows via CLI. Use when creating, reading,
  or modifying Excel workbooks from scripts, CI/CD, or coding agents.
  Supports Power Query, DAX, PivotTables, Tables, Ranges, Charts, VBA.
  Triggers: Excel, spreadsheet, workbook, xlsx, excelcli, CLI automation.
compatibility: Windows + Microsoft Excel 2016+ required. Uses COM interop - does NOT work on macOS or Linux.
allowed-tools: Cmd(excelcli:*),PowerShell(excelcli:*)
disable-model-invocation: true
license: MIT
version: 1.4.0
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://excelmcpserver.dev/
---

# Excel Automation with excelcli

## CRITICAL RULES (MUST FOLLOW)

### Rule 1: NEVER Ask Clarifying Questions

Execute commands to discover the answer instead:

| DON'T ASK | DO THIS INSTEAD |
|-----------|-----------------|
| "Which file should I use?" | `excelcli -q session list` |
| "What table should I use?" | `excelcli -q table list --session <id>` |
| "Which sheet has the data?" | `excelcli -q sheet list --session <id>` |

**You have commands to answer your own questions. USE THEM.**

### Rule 2: Use File-Based Input for Complex Data

PowerShell cannot reliably parse inline JSON or M code:

```powershell
# CORRECT - Use --values-file
$data = '[["Name","Price"],["Widget",100]]'
$data | Out-File -Encoding UTF8 C:\temp\data.json
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1 --values-file C:\temp\data.json

# CORRECT - Use --mcode-file for Power Query
$mcode | Out-File -Encoding UTF8 C:\temp\query.m
excelcli -q powerquery create --session 1 --query Products --mcode-file C:\temp\query.m
```

### Rule 3: Session Lifecycle

```powershell
excelcli -q session open C:\path\file.xlsx     # Returns session ID
excelcli -q range set-values --session 1 ...   # Use session ID
excelcli -q session close --session 1 --save   # Save and release
```

**Unclosed sessions leave Excel processes running, locking files.**

### Rule 4: Data Model Prerequisites

DAX operations require tables in the Data Model:

```powershell
excelcli -q table add-to-datamodel --session 1 --table Sales  # Step 1
excelcli -q datamodel create-measure --session 1 ...          # Step 2 - NOW works
```

### Rule 5: Power Query Development Lifecycle

**BEST PRACTICE: Test M code before creating permanent queries**

```powershell
# Step 1: Test M code without persisting (catches errors early)
excelcli -q powerquery evaluate --session 1 --mcode-file query.m

# Step 2: Create permanent query with validated code
excelcli -q powerquery create --session 1 --query Q1 --mcode-file query.m

# Step 3: Load data to destination
excelcli -q powerquery refresh --session 1 --query Q1
```

**Why evaluate first:** Better error messages than COM exceptions, see actual data preview, avoids broken queries in workbook.

### Rule 6: Report File Errors Immediately

If you see "File not found" or "Path not found" - STOP and report to user. Don't retry.

## Quick Reference

| Task | Command |
|------|---------|
| Open workbook | `excelcli -q session open <path>` |
| List sheets | `excelcli -q sheet list --session <id>` |
| Read data | `excelcli -q range get-values --session <id> --sheet Sheet1 --range A1:D10` |
| Create table | `excelcli -q table create --session <id> --sheet Sheet1 --range A1:D10 --name Sales` |
| Add to Data Model | `excelcli -q table add-to-datamodel --session <id> --table Sales` |
| Create pivot | `excelcli -q pivottable create --session <id> --source Sales --dest-sheet Analysis` |
| Save & close | `excelcli -q session close --session <id> --save` |

**Primary source:** `excelcli <command> --help` for authoritative parameter info.

## Reference Documentation

- @references/behavioral-rules.md - Core execution rules and LLM guidelines
- @references/anti-patterns.md - Common mistakes to avoid
- @references/workflows.md - Data Model constraints and patterns

## Installation

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

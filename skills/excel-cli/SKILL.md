---
name: excel-cli
description: >
  Automate Microsoft Excel on Windows via CLI. Use when creating, reading, 
  or modifying Excel workbooks from scripts, CI/CD, or coding agents.
  Supports Power Query, DAX, PivotTables, Tables, Ranges, Charts, VBA.
  Triggers: Excel, spreadsheet, workbook, xlsx, excelcli, CLI automation.
allowed-tools: Cmd(excelcli:*),PowerShell(excelcli:*)
disable-model-invocation: true
license: MIT
version: 1.4.0
tags:
  - excel
  - cli
  - automation
  - windows
  - powerquery
  - dax
  - scripting
repository: https://github.com/sbroenne/mcp-server-excel
documentation: https://excelmcpserver.dev/
---

# Excel Automation with excelcli

## ⚠️ CRITICAL RULES (READ FIRST)

### Rule 1: NEVER Ask Clarifying Questions

**STOP.** If you're about to ask a question, DON'T. Execute commands to discover the answer instead.

| ❌ DON'T ASK | ✅ DO THIS INSTEAD |
|--------------|-------------------|
| "Which file should I use?" | `excelcli -q session list` |
| "What table should I use?" | `excelcli -q table list --session <id>` |
| "Which sheet has the data?" | `excelcli -q sheet list --session <id>` |
| "What are the column names?" | `excelcli -q range get-values --session <id> --sheet Sheet1 --range A1:Z1` |

**You have commands to answer your own questions. USE THEM.**

### Rule 2: Report File/Path Errors Immediately

**If you see "File not found", "Path not found", or "Could not find" errors - STOP and report them.**

Don't try to work around missing files. Report: "Error: The file [path] does not exist."

These are infrastructure problems the user must fix, not something you can solve.

### Rule 3: Use File-Based Input for Complex Data (CRITICAL)

**PowerShell CANNOT reliably parse inline JSON arrays or M code with special keywords.**

**For `--values` with multi-row data:** Use `--values-file`
```powershell
# ✅ CORRECT - Create JSON file, then load it
$data = '[["Name","Price"],["Widget",100],["Gadget",200]]'
$data | Out-File -Encoding UTF8 C:\temp\data.json
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1 --values-file C:\temp\data.json

# ❌ WRONG - Inline JSON fails on commas
--values '[["Name","Price"],["Widget",100]]'  # PowerShell splits on commas!
```

**For `--mcode` with Power Query M code:** Use `--mcode-file`
```powershell
# ✅ CORRECT - Create .m file, then load it  
$mcode = 'let Source = Csv.Document(File.Contents("C:\data.csv")) in Source'
$mcode | Out-File -Encoding UTF8 C:\temp\query.m
excelcli -q powerquery create --session 1 --query Products --mcode-file C:\temp\query.m

# ❌ WRONG - The "in" keyword is a PowerShell reserved word!
--mcode "let Source = ... in Source"  # "in" parsed as PowerShell keyword!
```

**KEY INSIGHT:** If you see "Could not match 'X' with an argument" - USE FILE-BASED INPUT.

### Rule 4: Always Use -q Flag

```powershell
excelcli -q <command>  # Outputs clean JSON, no banner
```

### Rule 5: Session Lifecycle

```powershell
# 1. Create/open session
excelcli -q session create C:\path\file.xlsx   # New file
excelcli -q session open C:\path\file.xlsx     # Existing file

# 2. Work with session ID
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1 --values '[["Data"]]'

# 3. Close and save
excelcli -q session close --session 1 --save
```

---

## Quick Reference

### Core Commands

| Task | Command |
|------|---------|
| Create file | `excelcli -q session create <path>` |
| Open file | `excelcli -q session open <path>` |
| List sessions | `excelcli -q session list` |
| Close & save | `excelcli -q session close --session <id> --save` |
| List sheets | `excelcli -q sheet list --session <id>` |
| List tables | `excelcli -q table list --session <id>` |
| Get data | `excelcli -q range get-values --session <id> --sheet <name> --range A1:D10` |
| Set data | `excelcli -q range set-values --session <id> --sheet <name> --range A1 --values '[...]'` |

### Discover Commands (ALWAYS USE THIS)

**Don't memorize commands. Discover them dynamically:**

```powershell
excelcli actions              # List ALL commands and actions (40+)
excelcli actions range        # List actions for 'range' command
excelcli range --help         # Show ALL parameters with descriptions
excelcli range set-values --help  # Show parameters for specific action
```

**Example - find how to create a PivotTable:**
```powershell
excelcli actions | Select-String pivot    # Find pivot-related commands
excelcli pivottable --help                 # See all pivottable actions
excelcli pivottable create-from-datamodel --help  # Get exact parameters
```

**The CLI help is the authoritative source.** If you're unsure about parameters, run `--help`.

---

## Common Workflows

### Write Data to Excel

```powershell
excelcli -q session create C:\Reports\Sales.xlsx
# Use file-based input for multi-row data
$data = '[["Product","Q1","Q2"],["Widget",100,150],["Gadget",80,90],["Device",200,180]]'
$data | Out-File -Encoding UTF8 C:\temp\sales.json
excelcli -q range set-values --session 1 --sheet Sheet1 --range A1:C4 --values-file C:\temp\sales.json
excelcli -q table create --session 1 --sheet Sheet1 --range A1:C4 --table "SalesData" --has-headers
excelcli -q session close --session 1 --save
```

### Import CSV with Power Query

```powershell
excelcli -q session create C:\Reports\Analysis.xlsx
# Use file-based input for M code (contains "in" keyword)
$mcode = 'let Source = Csv.Document(File.Contents("C:\Data\sales.csv"), [Delimiter=","]) in Source'
$mcode | Out-File -Encoding UTF8 C:\temp\query.m
excelcli -q powerquery create --session 1 --query SalesData --mcode-file C:\temp\query.m
excelcli -q session close --session 1 --save
```

### Read and Analyze Data

```powershell
excelcli -q session open C:\Reports\Sales.xlsx
excelcli -q table list --session 1
excelcli -q table read --session 1 --table SalesData
excelcli -q session close --session 1
```

### Create PivotTable with DAX

```powershell
excelcli -q session open C:\Reports\Sales.xlsx
excelcli -q table add-to-datamodel --session 1 --table SalesData
excelcli -q datamodel create-measure --session 1 --table SalesData --measure "TotalSales" --expression "SUM(SalesData[Q1])"
excelcli -q pivottable create-from-datamodel --session 1 --table SalesData --dest-sheet Analysis --dest-cell A1 --pivot-table SalesPivot
excelcli -q session close --session 1 --save
```

---

## Error Recovery

| Error | Solution |
|-------|----------|
| "Could not match 'X' with an argument" | **USE FILE-BASED INPUT** - `--values-file` or `--mcode-file` |
| "Could not match 'in' with an argument" | M code `in` keyword issue - use `--mcode-file` |
| "File not found" / "Could not find" | **STOP** - Report error to user, don't retry |
| "Table not found" | Run `table list` to discover actual table names |
| "Sheet not found" | Run `sheet list` to discover sheet names |
| "Session not found" | Run `session list` to get valid session IDs |
| "Table not in Data Model" | Run `table add-to-datamodel` first |

---

## Reference Documentation

**Primary source:** Use `excelcli <command> --help` for authoritative, up-to-date parameter info.

For behavioral guidance and patterns, see:

- @references/behavioral-rules.md - Execution rules (format cells, use Tables)
- @references/anti-patterns.md - Common mistakes to avoid
- @references/workflows.md - Production workflow patterns

---

## Installation

```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

## Requirements

- Windows with Microsoft Excel 2016+
- .NET 10 Runtime
- VBA operations require "Trust access to VBA project object model" enabled

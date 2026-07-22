---
title: Troubleshooting & FAQ
description: >-
  Fixes for the most common Excel MCP Server issues — Excel must be closed,
  VBA trust, DAX/MSOLAP setup, PATH problems, and protected workbooks.
keywords: "Excel MCP troubleshooting, VBA trust, MSOLAP DAX, workbook locked, mcp-excel not recognized, IRM AIP Excel"
---

# Troubleshooting & FAQ

Hitting a snag? Most first-time issues fall into one of the cases below. If none
of these help, open a [GitHub issue](https://github.com/sbroenne/mcp-server-excel/issues).

## Frequently asked questions

??? question "Do I need to know how Excel automation works to use this?"
    No. You talk to your AI assistant in plain language ("build a PivotTable of
    sales by product and chart it") and it drives Excel for you. The
    [feature reference](features.md) is there when you want to see everything
    that's possible — you don't need to memorize it.

??? question "Does it require Microsoft Excel to be installed?"
    Yes. Excel MCP Server drives the **real Excel application** through its COM
    API, so it's **Windows-only** and needs **Excel 2016 or later** installed
    locally. It is not a file-format parser and does not run on macOS or Linux.

??? question "Will it damage my existing workbooks?"
    No. Excel itself opens and saves the file, so formulas, PivotTables, charts,
    macros, the Data Model, and formatting are all preserved. Other tools that
    rewrite the `.xlsx` file directly can silently drop those; here Excel does
    the work.

??? question "CLI or MCP Server — which should I install?"
    Both expose the **same 234 operations**. Use the **MCP Server** for
    conversational AI (Claude Desktop, VS Code Chat); use the **CLI**
    (`excelcli`) for coding agents and scripting, where it uses ~64% fewer
    tokens. You can install both. See [Installation](installation.md).

??? question "Does it cost anything or send my data anywhere?"
    Excel MCP Server is free and open source (MIT). It runs locally against your
    own Excel. A few opt-in features reach the internet (remote M/DAX
    formatting, and Python in Excel, which runs in Microsoft's cloud). See
    [Privacy](privacy.md) for details.

## Common issues

### "Workbook is locked" or "Cannot open file"

Close **all** open Excel windows before running Excel MCP Server. It needs
exclusive access to the workbook (an Excel COM limitation), so a file that's
already open in Excel can't be opened for automation.

### `mcp-excel` / `excelcli` is not recognized

The executable isn't on your `PATH`.

```powershell
# Confirm where it is (if anywhere)
where.exe mcp-excel
where.exe excelcli
```

Either add the folder containing the `.exe` to your `PATH` (see the
[MCP Server](installation-mcp-server.md) or [CLI](installation-cli.md)
installation guide), or use the full path in your MCP client config, e.g.
`"command": "C:\\Tools\\ExcelMcp\\mcp-excel.exe"`.

### VBA commands fail: "Programmatic access to Visual Basic Project is not trusted"

VBA operations need one manual Excel setting turned on:

1. Open Excel → **File → Options → Trust Center**
2. Click **Trust Center Settings**
3. Select **Macro Settings**
4. Check **"Trust access to the VBA project object model"**
5. Click **OK** twice

This is a Windows security setting — Excel MCP Server never changes it for you.
Also remember VBA lives in **`.xlsm`** workbooks, not `.xlsx`.

### DAX queries fail (`evaluate`, `execute-dmv`)

DAX query execution needs the **Microsoft Analysis Services OLE DB Provider
(MSOLAP)**, which isn't always installed with Office.

- **Easiest:** install [Power BI Desktop](https://powerbi.microsoft.com/desktop) (it includes MSOLAP).
- **Alternative:** install the [OLE DB Driver for Analysis Services](https://learn.microsoft.com/analysis-services/client-libraries).

### Protected (IRM / AIP) workbooks won't open

Rights-managed files need Excel visible so the sign-in or policy prompt can
appear. Keep Excel on screen while opening:

```powershell
excelcli session open "D:\Docs\Protected.xlsx" --show --timeout 120
```

With the MCP Server, ask your assistant to *"show me Excel while you work"* so
the authentication prompt is interactable. These files are opened read-only.

### Changes aren't taking effect / old version still running

Fully restart your MCP client (close VS Code or Claude Desktop completely,
including any background windows, then reopen). MCP servers are launched by the
client, so a stale process can linger until you restart it.

```powershell
# Confirm which version you're on
mcp-excel --version
excelcli --version
```

### `npx` commands fail

Auto-configuration (`add-mcp`) and skill installation use `npx`, which needs
**Node.js**:

```powershell
winget install OpenJS.NodeJS.LTS
```

## Still stuck?

- **Installation details:** [MCP Server](installation-mcp-server.md) · [CLI](installation-cli.md)
- **How it works:** [Architecture](architecture.md)
- **Report a bug or ask a question:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)

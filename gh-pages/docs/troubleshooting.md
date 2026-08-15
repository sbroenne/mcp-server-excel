---
title: Troubleshooting
description: >-
  Fixes for the most common Excel MCP Server issues - Excel must be closed,
  VBA trust, DAX/MSOLAP setup, PATH problems, and protected workbooks.
keywords: "Excel MCP troubleshooting, VBA trust, MSOLAP DAX, workbook locked, mcp-excel not recognized, IRM AIP Excel"
---

# Troubleshooting

Hitting a snag? Most first-time issues fall into one of the cases below. For
general questions about what the tool is and what it needs, see the
[FAQ](faq.md). If none of these help, open a
[GitHub issue](https://github.com/sbroenne/mcp-server-excel/issues).

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

- **General questions:** [FAQ](faq.md)
- **Task guides:** [Refresh Power Query](guides/refresh-power-query.md) · [PivotTables](guides/automate-pivottables.md) · [DAX & the Data Model](guides/query-data-model-with-dax.md) · [VBA macros](guides/run-vba-macros.md)
- **Installation details:** [MCP Server](installation-mcp-server.md) · [CLI](installation-cli.md)
- **How it works:** [Architecture](architecture.md)
- **Report a bug or ask a question:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)

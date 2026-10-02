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

For npx installations, bare `mcp-excel` / `excelcli` is not installed on PATH.
Use `npx -y @sbroenne/mcp-server-excel@latest` or
`npx -y @sbroenne/excelcli@latest` instead. For standalone or global
installations, check PATH:

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

- **Easiest:** install [Power BI Desktop](https://www.microsoft.com/en-us/power-platform/products/power-bi/desktop) (it includes MSOLAP).
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

Finish work and explicitly save/close the intended workbook sessions first.
For an npx-based MCP server, restart the server/client: `@latest` is resolved
at launch using normal npm caching. Standalone executables and older binary
MCPBs do not update just because the client restarts; replace the executable or
install the new npx-based MCPB.

```powershell
# Check the npm-launched executables
npx -y @sbroenne/mcp-server-excel@latest --version
npx -y @sbroenne/excelcli@latest --version
```

The CLI version command reports its foreground executable, not necessarily
the active background service. After safely closing workbook sessions, use
`npx -y @sbroenne/excelcli@latest -q service stop`; the next workbook command
starts the service from the selected CLI version. See the
[CLI update instructions](installation-cli.md#updating-the-cli) before stopping it.

### `npx` commands fail

The npm server/CLI, npx-based MCPB, auto-configuration (`add-mcp`), and skill
installation require **Node.js with npm/npx on PATH**. Claude's built-in Node.js
does not guarantee the external npx command is available:

```powershell
winget install OpenJS.NodeJS.LTS
```

Restart the client after installation so it receives the new PATH. Package
downloads and update checks need network access; npm's normal cache policy
still applies.
## Still stuck?

- **General questions:** [FAQ](faq.md)
- **Task guides:** [Refresh Power Query](guides/refresh-power-query.md) · [PivotTables](guides/automate-pivottables.md) · [DAX & the Data Model](guides/query-data-model-with-dax.md) · [VBA macros](guides/run-vba-macros.md)
- **Installation details:** [MCP Server](installation-mcp-server.md) · [CLI](installation-cli.md)
- **How it works:** [Architecture](architecture.md)
- **Report a bug or ask a question:** [GitHub Issues](https://github.com/sbroenne/mcp-server-excel/issues)

---
name: excel-cli
description: >-
  Discover and launch the Excel CLI for ordinary Excel workbook requests on
  Windows, including reads, writes, formulas, data refreshes, and workbook
  automation. Use the installed excel-cli plugin's npx launcher and native help
  to find commands; report presentation conventions are a separate optional skill.
compatibility: Windows with desktop Excel, Node.js 18 or later, and the excel-cli plugin.
---

# Excel CLI discovery

The plugin does not place `excelcli` on PATH. Its launcher runs
`npx -y @sbroenne/excelcli@latest`; do not assume a standalone executable is
installed or restore the retired GitHub-release downloader.

For the plugin-bundled copy of this skill, `..\..\bin\start-cli.ps1` relative to
the directory containing this `SKILL.md` locates the plugin launcher.
For a standalone skill installation (including the skill ZIP), that relative
path does not apply: the skill package does not contain the launcher. Locate
the actual installed `excel-cli` plugin through the client's plugin listing
and installation directory instead; do not resolve the launcher relative to
the standalone skill. If the required plugin is absent, report that prerequisite
rather than invoking a nonexistent file.
Replace `<plugin-root>` below with that plugin's actual installed directory:

```powershell
& "<plugin-root>\bin\start-cli.ps1" --help
& "<plugin-root>\bin\start-cli.ps1" session --help
```

Use the same full launcher path for subsequent commands, especially quoted JSON
arguments in Windows PowerShell. Discover exact command names, flags, defaults,
and save/close behavior through native `--help`; do not invent syntax.

Inspect the existing workbook state relevant to the request. Preserve unrelated
contents and save only the authorized result. If an operation fails, inspect its
reported partial state rather than assuming rollback or automatically cleaning up.
Load `excel-cli-report-formatting` only for requested report presentation, not
ordinary Excel reads or targeted updates.

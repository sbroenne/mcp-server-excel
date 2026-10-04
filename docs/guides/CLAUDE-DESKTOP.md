# Claude Desktop Configuration

Connect Claude Desktop to installed Microsoft Excel on Windows with Excel MCP
Server. Ask Claude to read workbook data, update cells, refresh Power Query, or
work with PivotTables through the real Excel application. Use the MCPB bundle
or a manual stdio configuration below.

## Requirements

- Windows 10 or later
- Microsoft Excel 2016 or later (desktop version)
- An interactive Windows desktop with Excel available to the signed-in user
- Node.js 18+ with npm/npx on PATH for MCPB or manual npx configuration
- Network access for npm package downloads and update checks

The published Windows packages are self-contained; no .NET runtime is required.

## Recommended: MCPB Bundle

1. Download `excel-mcp-{version}.mcpb` from the
   [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest).
2. Double-click the bundle or drag it into Claude Desktop.
3. Restart Claude Desktop.

The bundle configures Claude Desktop to run
`npx -y @sbroenne/mcp-server-excel@latest`. It contains configuration and
documentation, not a fixed server executable or bundled npm. Install Node.js
LTS first if npx is unavailable, then restart Claude Desktop so it sees PATH.
Do not assume Claude's built-in Node.js provides the external npx command.

`@latest` uses normal npm resolution and caching, not a guaranteed online
check on every launch. A running server does not change versions. Users of an
older binary MCPB must install the new bundle once; later bundle configuration
changes still require manual replacement.

## Manual Configuration

Add the server to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@sbroenne/mcp-server-excel@latest"]
    }
  }
}
```

Restart Claude Desktop after saving the configuration.

### Standalone Alternative

This method does not require Node.js, but executable updates are manual.

1. Download `ExcelMcp-MCP-Server-{version}-windows.zip` from the
   [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest).
2. Extract it to a permanent directory such as `C:\Tools\ExcelMcp`.
3. Add the server to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "C:\\Tools\\ExcelMcp\\mcp-excel.exe",
      "args": []
    }
  }
}
```

Restart Claude Desktop after saving the configuration.

## Try Your First Excel Request

After restarting Claude Desktop, try a read-only request using the full path
of an existing workbook. Replace this example path with your workbook's path:

```text
Open C:\Work\sales.xlsx and list its worksheet names. Do not change or save
the workbook.
```

Approve server or tool access if Claude prompts you. A response listing the
workbook's actual worksheets confirms that Claude can access desktop Excel,
not just read a spreadsheet file. For more tasks, see
[Excel automation examples](../USE-CASES.md).

## Recommended Workflow

```text
1. Discover existing sessions and confirm the user's intended full workbook path.
2. Create a new workbook or open that existing path.

3. Use the returned session ID for workbook operations.

4. Check results and save the intended successful changes. Close only when
   authorized and the session reports canClose: true.
```

Use full Windows paths and close sessions explicitly so Excel processes do not
remain open and lock workbooks.

## Troubleshooting

### Excel not found

- Confirm that desktop Excel 2016 or later is installed.
- Confirm that Excel starts normally for the current Windows user.

### Access denied or file locked

- Confirm that the path is writable.
- Reuse the matching session when possible; do not close another user's window.
- Ask for a different destination only if the supplied one cannot be used.

### COM timeout

- Check whether Excel is displaying a modal dialog.
- Allow long-running refresh or calculation operations to finish.
- Inspect the surviving sessions and partial changes before retrying. Restarting
  can lose unsaved work or trigger saving during normal shutdown.

### VBA operations fail

Read the actual error. VBA project inspection/editing requires trusted project
access configured manually by the user. Running an existing macro does not
itself require that project access, though Excel's macro security still applies.
Do not change Trust Center settings automatically or assume every VBA error is
a trust failure.

See the current
[MCP Server installation guide](https://excelmcpserver.dev/installation-mcp-server/)
for other supported clients and setup methods.

# Claude Desktop Configuration

Excel MCP Server works with Claude Desktop on Windows through the MCPB bundle
or a manual stdio configuration.

## Requirements

- Windows 10 or later
- Microsoft Excel 2016 or later (desktop version)

The published Windows packages are self-contained; no .NET runtime is required.

## Recommended: MCPB Bundle

1. Download `excel-mcp-{version}.mcpb` from the
   [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest).
2. Double-click the bundle or drag it into Claude Desktop.
3. Restart Claude Desktop.

The bundle contains the MCP server and configures Claude Desktop automatically.

## Manual Configuration

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

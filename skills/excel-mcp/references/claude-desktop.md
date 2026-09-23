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
1. Create or open a workbook:
   file(action: 'create', path: 'C:\Users\Me\Documents\report.xlsx')

2. Use the returned session ID for workbook operations.

3. Save and close:
   file(action: 'close', session_id: '...', save: true)
```

Use full Windows paths and close sessions explicitly so Excel processes do not
remain open and lock workbooks.

## Troubleshooting

### Excel not found

- Confirm that desktop Excel 2016 or later is installed.
- Confirm that Excel starts normally for the current Windows user.

### Access denied or file locked

- Confirm that the path is writable.
- Close any other Excel instance that already has the workbook open.
- Try a workbook in the user's Documents directory.

### COM timeout

- Check whether Excel is displaying a modal dialog.
- Allow long-running refresh or calculation operations to finish.
- Restart Claude Desktop if the Excel process is unresponsive.

### VBA operations fail

VBA project editing requires explicit trust:

1. Open Excel Options.
2. Select **Trust Center** and then **Trust Center Settings**.
3. Enable **Trust access to the VBA project object model**.
4. Restart Excel MCP Server.

See the current
[MCP Server installation guide](https://excelmcpserver.dev/installation-mcp-server/)
for other supported clients and setup methods.

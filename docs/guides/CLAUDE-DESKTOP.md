# Claude Desktop Configuration

Connect Claude Desktop to installed Microsoft Excel on Windows with Excel MCP
Server. Ask Claude to read workbook data, update cells, refresh Power Query, or
work with PivotTables through the real Excel application. Use the MCPB bundle
or a manual stdio configuration below.

## Requirements

- Windows x64/ARM64 with desktop Excel 2016 or later, or
- Apple Silicon macOS with Excel for Mac 16.112 or later
- An interactive desktop session; Intel Macs and headless hosts are unsupported

The published Windows x64/ARM64 and Apple Silicon macOS packages are self-contained;
no .NET runtime is required.

## Recommended: MCPB Bundle

1. Download `excel-mcp-{version}-windows.mcpb` or
   `excel-mcp-{version}-macos-arm64.mcpb` from the
   [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest).
2. Double-click the bundle or drag it into Claude Desktop.
3. Restart Claude Desktop.

The bundle contains the MCP server and configures Claude Desktop automatically.

## Manual Configuration

1. Download `ExcelMcp-MCP-Server-{version}-windows.zip` or
   `ExcelMcp-MCP-Server-{version}-macos-arm64.zip` from the
   [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest).
2. Extract it to a permanent directory.
3. Add the matching absolute executable path to Claude's configuration:
   `%APPDATA%\Claude\claude_desktop_config.json` on Windows, or
   `~/Library/Application Support/Claude/claude_desktop_config.json` on Mac.

Windows example:

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

Mac example (replace the user and installation path):

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "/Users/your-user/Tools/ExcelMcp/mcp-excel",
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

Use full native paths. Windows sessions own their Excel processes; Mac sessions
own only exact workbooks in shared Excel. Never terminate shared Excel or close
unrelated workbooks.

## Troubleshooting

### Excel not found

- Confirm that the required desktop Excel version for your platform is installed.
- Confirm that Excel starts normally for the current user.

### Access denied or file locked

- Confirm that the path is writable.
- Reuse the matching session when possible; do not close another user's window.
- Ask for a different destination only if the supplied one cannot be used.

### Mac permission, platform, or recovery errors

- Grant Excel Automation permission manually when macOS requests it.
- `PlatformNotSupported` means the action/variant is unavailable; installing an
  optional bridge does not enable unaccepted actions.
- For `RecoveryRequired`, reconcile the exact workbook and pending dialogs
  manually before restarting the client. Do not repeat the open, close other
  workbooks, or kill shared Excel.

### COM timeout (Windows)

- Check whether Excel is displaying a modal dialog.
- Allow long-running refresh or calculation operations to finish.
- Inspect the surviving sessions and partial changes before retrying. Restarting
  can lose unsaved work or trigger saving during normal shutdown.

### VBA operations fail (Windows)

Read the actual error. VBA project inspection/editing requires trusted project
access configured manually by the user. Running an existing macro does not
itself require that project access, though Excel's macro security still applies.
Do not change Trust Center settings automatically or assume every VBA error is
a trust failure.

VBA is unsupported in the Mac beta; changing macro preferences cannot enable it.

See the current
[MCP Server installation guide](https://excelmcpserver.dev/installation-mcp-server/)
for other supported clients and setup methods.

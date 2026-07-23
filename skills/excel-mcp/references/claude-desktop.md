# Claude Desktop Configuration

Excel MCP Server works with Claude Desktop on Windows, but requires specific configuration for the Windows container environment.

## Configuration Location

Claude Desktop config file:
```
%APPDATA%\Claude\claude_desktop_config.json
```

## Basic Configuration

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "excel-mcp-server.exe",
      "args": []
    }
  }
}
```

Or using the .NET tool:
```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "dotnet",
      "args": ["excel-mcp-server"]
    }
  }
}
```

## Windows Container Considerations

Claude Desktop runs in a Windows container with specific constraints:

### File System Access

The container has limited file system access. Excel files should be in accessible locations:

- **User Documents**: `C:\Users\<username>\Documents\`
- **User Desktop**: `C:\Users\<username>\Desktop\`
- **Temp directory**: `%TEMP%` or `C:\Users\<username>\AppData\Local\Temp\`

**Recommendation**: Work with files in your Documents folder.

### Excel Instance

- Excel MCP Server manages its own Excel instance via COM automation
- The Excel window may be visible or hidden depending on operation
- Long-running operations show Excel's progress indicators

### Session Persistence

Sessions are tied to the Claude Desktop session:
- Closing Claude Desktop terminates active Excel sessions
- Unsaved changes may be lost
- Use explicit `file(action: 'close', save: true)` to persist work

## Recommended Workflow

```
1. Create or open file in accessible location:
   file(action: 'create', filePath: 'C:\\Users\\Me\\Documents\\report.xlsx')

2. Perform operations with returned sessionId

3. Explicitly save and close when done:
   file(action: 'close', sessionId: '...', save: true)
```

## Troubleshooting

### "Excel not found" Error
- Ensure Microsoft Excel is installed on the Windows system
- Excel 2016, 2019, 2021, or Microsoft 365 required

### "Access denied" Error
- Check file path is in accessible directory
- Ensure file is not open in another Excel instance
- Try using Documents folder instead of other locations

### "COM timeout" Error
- Excel may be showing a dialog - check for visible Excel window
- Operation may be long-running - wait for completion
- Restart Claude Desktop if Excel becomes unresponsive

### VBA Operations Fail
VBA requires explicit trust setting in Excel:
1. Open Excel Options → Trust Center → Trust Center Settings
2. Enable "Trust access to the VBA project object model"
3. Restart Excel MCP Server

## MCPB Bundle Alternative

For simplified installation, use the MCPB bundle which auto-configures Claude Desktop:

1. Download `excel-mcp-bundle.mcpb` from releases
2. Double-click to install
3. Restart Claude Desktop

See the main repository for MCPB installation instructions.

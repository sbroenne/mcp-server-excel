# Version Checking and Update Notifications

This document describes how version checking and update notifications work in ExcelMcp.

## Overview

ExcelMcp provides version checking in two contexts:

1. **CLI Tool** - Manual version check and automatic daemon startup notification
2. **MCP Server** - Protocol-level version negotiation (handled by MCP SDK)

## CLI Version Checking

### Manual Version Check

Users can check for updates at any time using the `--version` flag:

```powershell
excelcli --version
```

This command:
1. Displays current version information
2. Checks NuGet for the latest version (non-blocking, 5-second timeout)
3. Shows a friendly message if an update is available
4. Provides the exact command to update

**Example output when update is available:**
```
⚠ Update available: 1.0.0 → 1.1.0
Run: dotnet tool update --global Sbroenne.ExcelMcp.CLI
Release notes: https://github.com/sbroenne/mcp-server-excel/releases/latest
```

### Automatic Daemon Notification

When the CLI daemon starts, it automatically checks for updates in the background:

1. **Timing**: Check occurs 5 seconds after daemon startup
2. **Non-blocking**: Version check runs asynchronously and never blocks daemon operations
3. **Silent on failure**: If the check fails (network error, timeout), no notification is shown
4. **Windows notification**: If an update is available, a system tray notification appears

**Notification Details:**
- **Title**: "Excel CLI Update Available"
- **Message**: Shows current version, new version, and update command
- **Duration**: 3 seconds (Windows standard)
- **Type**: Info balloon (NotifyIcon.ShowBalloonTip)

### Implementation Details

**Components:**

1. **`DaemonVersionChecker.cs`** - Core version checking logic
   - `CheckForUpdateAsync()` - Compares current version with latest NuGet version
   - Returns `UpdateInfo` if update available, `null` otherwise
   - Non-blocking with 5-second timeout (inherited from `NuGetVersionChecker`)

2. **`DaemonTray.cs`** - System tray UI
   - `ShowUpdateNotification()` - Displays Windows notification
   - Thread-safe (invokes on UI thread if needed)
   - Integrates with existing tray icon

3. **`ExcelDaemon.cs`** - Daemon startup
   - Triggers version check 5 seconds after startup
   - Runs in background Task.Run() to avoid blocking
   - Fails silently on any errors

**Best Practices Followed:**

1. **Non-intrusive**: Balloon tip notification, not modal dialog
2. **Non-blocking**: Runs asynchronously after daemon is fully initialized
3. **Fail-safe**: All errors caught and ignored - never impacts daemon operation
4. **Windows-native**: Uses NotifyIcon balloon tips (Windows Forms standard)
5. **Actionable**: Message includes exact command to update

### Future Enhancements (Optional)

The current implementation uses classic balloon tips. For Windows 11, consider upgrading to modern toast notifications:

1. **Package**: `Microsoft.Toolkit.Uwp.Notifications`
2. **Benefits**: Richer content, action buttons, Windows 11 Action Center integration
3. **Migration**: Keep balloon tip as fallback for older Windows versions

## MCP Server Version Handling

### Protocol Version Negotiation

The MCP Server uses the official Model Context Protocol version negotiation mechanism. **No custom implementation is needed** for application version checking.

**How it works:**

1. **Initialization Handshake**: Client and server negotiate protocol version during `initialize` request
2. **Version Format**: `YYYY-MM-DD` format (e.g., "2025-11-25")
3. **Compatibility**: Both parties must agree on a mutually supported protocol version
4. **Error Handling**: If no compatible version, server returns JSON-RPC error (-32001)

**Server Info:**

The MCP Server includes application version in the `ServerInfo` response:

```json
{
  "name": "excel-mcp",
  "version": "1.0.0"
}
```

This is sent during the initialization handshake and visible to MCP clients.

### Why No Custom Version Check?

The MCP specification already provides:

1. **Protocol Version Negotiation**: Ensures client/server compatibility
2. **ServerInfo Exchange**: Makes application version visible to clients
3. **Standard Error Codes**: Well-defined behavior for version mismatches

Adding a separate version check mechanism would:
- Duplicate functionality already in the protocol
- Create confusion between protocol version and application version
- Not follow MCP best practices

### MCP Spec References

- **Versioning**: [MCP Specification - Versioning](https://modelcontextprotocol.io/specification/versioning)
- **Initialization**: Protocol version sent in `initialize` request params
- **Error Codes**: -32001 for protocol version mismatch

## Testing

### Unit Tests

**Location**: `tests/ExcelMcp.CLI.Tests/Unit/DaemonVersionCheckerTests.cs`

Tests verify:
1. Version comparison logic
2. Graceful failure handling (network errors, timeouts)
3. UpdateInfo message formatting
4. Notification message content

**Run tests:**
```powershell
dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj --filter "Feature=VersionCheck"
```

### Manual Testing

**Test daemon notification:**
1. Start daemon: `excelcli daemon start`
2. Wait 5 seconds after startup
3. If update is available, Windows notification should appear in system tray

**Test version flag:**
1. Run: `excelcli --version`
2. Verify output shows current version and checks NuGet
3. If update available, message includes update command

## Configuration

Currently, version checking is enabled by default with no configuration options.

**Future options could include:**
- Disable version check entirely
- Adjust check frequency for daemon
- Opt-out of notifications (check only on manual request)

These would require adding configuration to `DaemonVersionChecker` or daemon settings.

## Troubleshooting

**No notification shown:**
- Check: Is an update actually available? Run `excelcli --version` to verify
- Check: Network connectivity (version check requires internet to reach NuGet)
- Check: Daemon logs for any errors during version check

**Update command fails:**
- Ensure you have internet connectivity
- Verify NuGet package manager is working: `dotnet tool list --global`
- Try updating manually: `dotnet tool update --global Sbroenne.ExcelMcp.CLI`

**Version check takes too long:**
- Timeout is 5 seconds by default (from `NuGetVersionChecker`)
- If network is slow, check may fail silently - this is intentional to avoid blocking

## References

- **Windows Notifications Best Practices**: [Microsoft Learn - Toast Notifications](https://learn.microsoft.com/en-us/windows/apps/develop/notifications/app-notifications/send-local-toast)
- **MCP Protocol Versioning**: [MCP Specification](https://modelcontextprotocol.io/specification/versioning)
- **NuGet API**: Version check uses NuGet v3 API at `https://api.nuget.org/v3-flatcontainer/`

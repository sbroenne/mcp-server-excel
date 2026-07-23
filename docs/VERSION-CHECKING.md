# Version Checking and Update Notifications

This document describes how version checking and update notifications work in ExcelMcp.

## Overview

ExcelMcp checks for updates against the latest GitHub Release. Both the CLI and MCP Server
use the GitHub Releases API to compare the current version with the latest published release.

## CLI Version Checking

### Manual Version Check

Users can check for updates at any time using the `--version` flag:

```powershell
excelcli --version
```

This command:
1. Displays current version information
2. Checks GitHub Releases for the latest version (non-blocking, 5-second timeout)
3. Shows a friendly message if an update is available
4. Provides the download URL for the new release

**Example output when update is available:**
```
⚠ Update available: 1.0.0 → 1.1.0
Download: https://github.com/sbroenne/mcp-server-excel/releases/latest
```

### Automatic Service Notification

When the ExcelMCP Service starts, it automatically checks for updates in the background:

1. **Timing**: Check occurs 5 seconds after service startup
2. **Non-blocking**: Version check runs asynchronously and never blocks service operations
3. **Silent on failure**: If the check fails (network error, timeout), no notification is shown
4. **Windows notification**: If an update is available, a system tray notification appears

**Notification Details:**
- **Title**: "Update Available"
- **Message**: Shows current version, new version, and download URL
- **Duration**: 3 seconds (Windows standard)
- **Type**: Info balloon (NotifyIcon.ShowBalloonTip)

## MCP Server Version Checking

The MCP Server checks for updates when `--version` is passed as a command-line argument:

```powershell
mcp-excel --version
```

**Example output when update is available:**
```
Excel MCP Server v1.0.0

Update available: 1.0.0 -> 1.1.0
Download: https://github.com/sbroenne/mcp-server-excel/releases/latest
```

## Implementation Details

### GitHub Releases API

Both CLI and MCP Server use the same endpoint:

```
GET https://api.github.com/repos/sbroenne/mcp-server-excel/releases/latest
```

**Response:**
```json
{
  "tag_name": "v1.2.3",
  ...
}
```

The `tag_name` field is stripped of the `v` prefix to get the version string: `v1.2.3` → `1.2.3`.

### Components

1. **`NuGetVersionChecker.cs`** (CLI and MCP Server) - GitHub Releases API client
   - `GetLatestVersionAsync()` - Fetches latest release tag from GitHub
   - Returns version string (e.g., `"1.2.3"`) or `null` if check failed
   - Non-blocking with 5-second timeout
   - Adds `User-Agent` header (required by GitHub API)

2. **`CliServiceTray.cs`** - System tray UI (CLI)
   - `CheckForUpdateAsync()` - Compares current version with latest GitHub release
   - `ShowBalloon()` - Displays Windows notification if update available

3. **`McpServerVersionChecker.cs`** (MCP Server) - Version comparison
   - `CheckForUpdateAsync()` - Compares current version with latest GitHub release
   - `GetCurrentVersion()` - Reads version from assembly metadata

### Best Practices Followed

1. **Non-intrusive**: Balloon tip notification or console output, not modal dialog
2. **Non-blocking**: Runs asynchronously after service is fully initialized
3. **Fail-safe**: All errors caught and ignored - never impacts service operation
4. **Windows-native**: Uses NotifyIcon balloon tips (Windows Forms standard)
5. **Actionable**: Message includes direct download URL

## How to Update

When a new version is available:

**Standalone exe (primary):**
1. Download the latest release from:
   - [https://github.com/sbroenne/mcp-server-excel/releases/latest](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract the new exe(s):
   - `ExcelMcp-MCP-Server-{version}-windows.zip` → `mcp-excel.exe`
   - `ExcelMcp-CLI-{version}-windows.zip` → `excelcli.exe`
3. Replace the existing exe(s) in your installation directory
4. Restart your MCP client

**NuGet (secondary):**
```powershell
dotnet tool update --global Sbroenne.ExcelMcp.McpServer
dotnet tool update --global Sbroenne.ExcelMcp.CLI
```

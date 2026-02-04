# Excel CLI Daemon Improvements - Implementation Summary

## Overview
This PR implements six improvements to the Excel CLI daemon's tray icon experience, making it more user-friendly and reducing the risk of data loss.

## Changes Made

### 1. ✅ Copyable Update Instructions
**Requirement:** Toast notification instructions should be copyable

**Implementation:** 
- Modified `ShowUpdateNotification()` in `DaemonTray.cs` to add text mentioning the "Update CLI" menu option
- Instead of just showing the update command in a balloon tip, we now:
  - Add a visible "Update CLI" menu item to the tray icon
  - Show the command in a confirmation dialog before execution
  - Allow users to copy the command from the dialog

**Result:** Users can now easily access and copy update instructions from both the notification and the confirmation dialog.

### 2. ✅ One-Click Update with Installation Detection
**Requirement:** Add "Update CLI" menu option when update is available, detect local vs global install

**Implementation:**
- Created `ToolInstallationDetector.cs` class:
  - `GetInstallationType()` - Detects if tool is installed globally or locally
  - `GetUpdateCommand()` - Returns appropriate update command based on install type
  - `TryUpdateAsync()` - Executes the update command and returns result
- Modified `DaemonTray.cs`:
  - Added `_updateMenuItem` field (initially hidden)
  - Modified `ShowUpdateNotification()` to store update info and show menu item
  - Added `UpdateCli()` method to show confirmation dialog and run update
  - Added `ShowUpdateResult()` method to show success/failure and restart daemon

**Result:** When an update is available, users see a menu option "Update to X.X.X" that:
1. Shows a confirmation dialog with the exact command that will run
2. Runs the update in the background
3. Shows success/failure result
4. Auto-restarts the daemon on success

### 3. ✅ Save Prompt When Closing Individual Sessions
**Requirement:** Ask user if they want to save when closing a session from the tray, with Cancel option

**Implementation:**
- Modified `RefreshSessionsMenu()` in `DaemonTray.cs`:
  - Removed separate "Close" and "Save & Close" menu items
  - Added single "Close Session..." menu item that calls `PromptCloseSession()`
- Added `PromptCloseSession()` method:
  - Shows MessageBox with Yes/No/Cancel buttons
  - Yes = save and close
  - No = close without saving
  - Cancel = abort operation

**Result:** When closing a session, users are prompted "Do you want to save changes to 'filename' before closing?" with three clear options.

### 4. ✅ Remove Greyed Out Menu Entry
**Requirement:** Remove the disabled "Excel CLI Daemon" menu entry

**Implementation:**
- Removed the following code from `DaemonTray.cs` constructor:
  ```csharp
  var statusItem = new ToolStripMenuItem("Excel CLI Daemon") { Enabled = false };
  _contextMenu.Items.Add(statusItem);
  ```

**Result:** The tray menu is now cleaner without the redundant disabled entry.

### 5. ✅ Save Prompt When Stopping Daemon
**Requirement:** "Stop Daemon" should ask user if they want to save sessions, with Cancel option

**Implementation:**
- Modified `StopDaemon()` in `DaemonTray.cs`:
  - Changed from Yes/No to Yes/No/Cancel dialog
  - Added explicit handling for each option:
    - Yes = save all sessions before stopping
    - No = close all sessions without saving
    - Cancel = abort stop operation
  - Added error handling with option to continue on save failure
  - Shows balloon tip with results

**Result:** When stopping the daemon with active sessions, users are prompted "Do you want to save all sessions before stopping the daemon?" with three options, preventing accidental data loss.

### 6. ✅ Update CLI Documentation
**Requirement:** Make sure the CLI doc is up-to-date

**Implementation:**
- Updated `src/ExcelMcp.CLI/README.md`:
  - Added "Auto-updates" feature to daemon description
  - Updated tray icon features section with new behaviors
  - Changed timeout from 5 to 10 minutes (matching actual code)
  - Updated descriptions to reflect Yes/No/Cancel dialogs
  - Added "Update CLI" menu feature with emoji
- Updated `CHANGELOG.md`:
  - Added entry for daemon improvements under [Unreleased]
  - Listed all new features and changes

**Result:** Documentation now accurately reflects the daemon's capabilities.

## Technical Details

### File Changes
1. **Created:** `src/ExcelMcp.CLI/Infrastructure/ToolInstallationDetector.cs` (108 lines)
   - Installation type detection
   - Update command generation
   - Update execution logic

2. **Modified:** `src/ExcelMcp.CLI/Daemon/DaemonTray.cs`
   - Added fields: `_updateMenuItem`, `_availableUpdate`
   - Modified methods: `ShowUpdateNotification()`
   - Added methods: `UpdateCli()`, `ShowUpdateResult()`, `PromptCloseSession()`
   - Modified methods: `RefreshSessionsMenu()`, `StopDaemon()`
   - Removed: Disabled status menu item

3. **Modified:** `src/ExcelMcp.CLI/README.md`
   - Updated daemon features section

4. **Modified:** `CHANGELOG.md`
   - Added unreleased changes entry

### User Experience Flow

#### Closing a Single Session
1. User right-clicks tray icon
2. Hovers over "Sessions (N)"
3. Clicks on a file → "Close Session..."
4. Dialog appears: "Do you want to save changes to 'filename' before closing?"
   - Yes → Session saved and closed
   - No → Session closed without saving
   - Cancel → Operation aborted, session remains open

#### Stopping Daemon with Active Sessions
1. User right-clicks tray icon → "Stop Daemon"
2. Dialog appears: "There are N active session(s). Do you want to save all sessions before stopping the daemon?"
   - Yes → All sessions saved and closed, daemon stops
   - No → All sessions closed without saving, daemon stops
   - Cancel → Operation aborted, daemon keeps running

#### Updating CLI
1. Daemon checks for updates on startup (after 5 seconds)
2. If update available:
   - Toast notification appears with message
   - "Update to X.X.X" menu item appears in tray
3. User clicks "Update to X.X.X"
4. Confirmation dialog shows the exact command that will run
   - OK → Update runs, success dialog appears, daemon restarts
   - Cancel → Operation aborted
5. If update fails, error dialog shows with manual instructions

## Testing Notes

Since we're running on Linux, we cannot build or test the Windows-specific code. The changes should be tested on Windows:

1. **Installation Detection:**
   - Test with global install: `dotnet tool install --global Sbroenne.ExcelMcp.CLI`
   - Test with local install: `dotnet tool install Sbroenne.ExcelMcp.CLI`
   - Verify correct update command is shown

2. **Session Close:**
   - Create multiple sessions
   - Try closing with Yes/No/Cancel
   - Verify files are saved/not saved as expected
   - Verify Cancel keeps session open

3. **Daemon Stop:**
   - Create multiple sessions
   - Stop daemon with Yes/No/Cancel
   - Verify all sessions handled correctly
   - Verify Cancel keeps daemon running

4. **Update Flow:**
   - Mock an old version to trigger update notification
   - Verify menu item appears
   - Verify confirmation dialog shows correct command
   - Test update success path
   - Test update failure path

5. **Documentation:**
   - Verify README accurately describes new features
   - Verify CHANGELOG entry is complete

## Potential Edge Cases

1. **Update while sessions active:** The update process restarts the daemon, which should close all sessions. Consider adding a warning about this.

2. **Network failure during update:** The TryUpdateAsync method catches exceptions and returns failure, which shows error dialog.

3. **Permission issues:** If user doesn't have permission to update (e.g., admin-installed global tool), the update will fail gracefully with error dialog.

4. **Concurrent operations:** All UI operations run on the Windows Forms UI thread via Invoke when needed, so thread safety should be maintained.

## Code Quality

- All methods follow existing patterns in the codebase
- Error handling with try-catch blocks
- Thread-safe UI updates using Invoke
- Clear user messaging with descriptive dialogs
- No breaking changes to existing functionality
- Backward compatible (old daemon behavior still works if update disabled)

## Future Enhancements

1. Could add option to "Save As" when closing sessions
2. Could add "Update and restart now" vs "Update on next restart" option
3. Could show changelog/release notes in update dialog
4. Could add progress bar for update process
5. Could remember user's preference for saving (Yes/No) across sessions

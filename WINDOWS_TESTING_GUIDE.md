# Windows Testing Guide for Daemon Improvements

## Overview
This guide provides step-by-step instructions for testing the daemon improvements on Windows. All code has been implemented and committed. Your task is to build, test, and verify the UI changes work correctly.

## Prerequisites
- Windows 10/11
- .NET 10 SDK installed
- Excel installed (COM automation requires Excel)
- Git repository cloned

## Step 1: Build the Project

```powershell
# Navigate to repository root
cd /path/to/mcp-server-excel

# Build the CLI project
dotnet build src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release

# Verify build succeeded with 0 warnings
# Expected: Build succeeded with 0 warnings, 0 errors
```

**Checkpoint:** Build must complete with 0 warnings before proceeding.

## Step 2: Install CLI Globally (for testing)

```powershell
# Uninstall existing version (if any)
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI

# Pack the NuGet package
dotnet pack src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release -o ./nupkg

# Install from local package
dotnet tool install --global Sbroenne.ExcelMcp.CLI --add-source ./nupkg --version <version-from-csproj>

# Verify installation
excelcli --version
```

## Step 3: Create Test Files

```powershell
# Create a test directory
mkdir C:\temp\excelcli-test
cd C:\temp\excelcli-test

# Create test Excel files (the CLI will create them)
# We'll create them via CLI commands in the next steps
```

## Step 4: Start Daemon and Create Sessions

```powershell
# Create first test file with session
excelcli session open test1.xlsx

# Expected output: JSON with sessionId
# Save the sessionId for later use

# Create second test file with session
excelcli session open test2.xlsx

# Expected output: JSON with sessionId
# Save this sessionId too

# Verify daemon is running
# Look for Excel CLI icon in Windows system tray (bottom-right)
```

**Checkpoint:** System tray should show Excel CLI icon. Verify icon is visible.

## Step 5: Test Session Close with Save Prompt

### Test 5.1: Close with Yes (Save)
1. Right-click the Excel CLI tray icon
2. Hover over "Sessions (2)" to see submenu
3. You should see two files: `test1.xlsx` and `test2.xlsx`
4. Click on `test1.xlsx` → "Close Session..."
5. Dialog should appear: "Do you want to save changes to 'test1.xlsx' before closing?"
6. **TAKE SCREENSHOT** of this dialog
7. Click "Yes"
8. Verify balloon tip appears: "Session Closed - Session saved and closed."
9. Right-click tray again, verify "Sessions (1)" now shows only `test2.xlsx`

### Test 5.2: Close with No (Don't Save)
1. Make a change to test2.xlsx (use CLI or open in Excel)
2. Right-click tray icon → "Sessions (1)" → `test2.xlsx` → "Close Session..."
3. Dialog should appear with Yes/No/Cancel buttons
4. **TAKE SCREENSHOT** of the dialog
5. Click "No"
6. Verify balloon tip: "Session closed without saving."
7. Right-click tray, verify "Sessions (0)" shows "No active sessions"

### Test 5.3: Close with Cancel
1. Create new session: `excelcli session open test3.xlsx`
2. Right-click tray → "Sessions (1)" → `test3.xlsx` → "Close Session..."
3. Click "Cancel" in the dialog
4. Verify session is still open (tray shows "Sessions (1)")

**Checkpoint:** All three buttons (Yes/No/Cancel) must work correctly. Screenshots required for Yes and No dialogs.

## Step 6: Test Daemon Stop with Save Prompt

### Test 6.1: Stop with Active Sessions - Yes (Save All)
1. Create multiple sessions:
   ```powershell
   excelcli session open test4.xlsx
   excelcli session open test5.xlsx
   ```
2. Right-click tray icon → "Stop Daemon"
3. Dialog should appear: "There are 2 active session(s). Do you want to save all sessions before stopping the daemon?"
4. **TAKE SCREENSHOT** of this dialog (should show Yes/No/Cancel buttons)
5. Click "Yes"
6. Verify balloon tip: "Sessions Saved - Saved and closed 2 session(s)."
7. Verify daemon stops (tray icon disappears)

### Test 6.2: Stop with Active Sessions - No (Don't Save)
1. Start daemon again (create a session): `excelcli session open test6.xlsx`
2. Verify tray icon appears
3. Right-click tray → "Stop Daemon"
4. Dialog should show Yes/No/Cancel
5. Click "No"
6. Verify daemon stops (tray icon disappears)
7. Sessions should be closed without saving

### Test 6.3: Stop with Active Sessions - Cancel
1. Start daemon: `excelcli session open test7.xlsx`
2. Right-click tray → "Stop Daemon"
3. Click "Cancel" in the dialog
4. Verify daemon keeps running (tray icon still visible)
5. Verify session still open: `excelcli session list`

### Test 6.4: Stop with No Sessions
1. Close all sessions manually
2. Right-click tray → "Stop Daemon"
3. Should stop immediately without dialog (no active sessions)
4. Verify daemon stops

**Checkpoint:** All stop scenarios must work correctly. Screenshot of the save prompt dialog required.

## Step 7: Test Menu Cleanup (Removed Item)

1. Start daemon: `excelcli session open test8.xlsx`
2. Right-click tray icon
3. **TAKE SCREENSHOT** of the entire context menu
4. Verify menu structure:
   ```
   Sessions (1)
   ─────────────
   Update to X.X.X  (only if update available)
   Stop Daemon
   ```
5. **VERIFY:** There should be NO disabled "Excel CLI Daemon" menu entry
6. Menu should be clean without redundant items

**Checkpoint:** Screenshot confirms no greyed-out "Excel CLI Daemon" entry.

## Step 8: Test Update Detection and Menu (If Possible)

### Option A: If Newer Version Exists on NuGet
1. Install an older version of the CLI
2. Start daemon: `excelcli session open test9.xlsx`
3. Wait 5 seconds for update check to complete
4. Toast notification should appear: "Excel CLI Update Available"
5. **TAKE SCREENSHOT** of the toast notification
6. Right-click tray icon
7. Verify "Update to X.X.X" menu item is visible
8. **TAKE SCREENSHOT** of menu showing update option

### Option B: Mock Update Test (Requires Code Modification)
If no newer version exists, you can temporarily modify the version check:
1. Edit `src/ExcelMcp.CLI/Infrastructure/DaemonVersionChecker.cs`
2. In `CheckForUpdateAsync`, hardcode return:
   ```csharp
   return new UpdateInfo
   {
       CurrentVersion = "1.0.0",
       LatestVersion = "2.0.0",
       UpdateAvailable = true
   };
   ```
3. Rebuild and reinstall CLI
4. Start daemon
5. Verify toast notification appears
6. Verify menu shows "Update to 2.0.0"
7. **TAKE SCREENSHOT** of menu

### Test Update Flow (Only if Update Available)
1. Click "Update to X.X.X" menu item
2. Confirmation dialog should appear showing:
   - Current and new versions
   - Exact command that will run
   - "The daemon will restart after update"
3. **TAKE SCREENSHOT** of confirmation dialog
4. Click "Cancel" (don't actually update during testing)
5. Verify operation cancelled

**Note:** For full update test, you would click "OK" and verify:
- Progress balloon appears
- Update executes
- Success/failure dialog shows
- Daemon restarts on success

## Step 9: Test Installation Type Detection

### Test 9.1: Global Installation
1. Verify CLI is installed globally (done in Step 2)
2. Check installation location:
   ```powershell
   (Get-Command excelcli).Source
   # Should be in %USERPROFILE%\.dotnet\tools
   ```
3. Start daemon, trigger update check (see Step 8)
4. Click update menu → verify dialog shows:
   ```
   dotnet tool update --global Sbroenne.ExcelMcp.CLI
   ```
5. **TAKE SCREENSHOT** showing global command

### Test 9.2: Local Installation (Optional)
1. Uninstall global: `dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI`
2. Install locally in a test project:
   ```powershell
   mkdir C:\temp\local-cli-test
   cd C:\temp\local-cli-test
   dotnet new tool-manifest
   dotnet tool install Sbroenne.ExcelMcp.CLI --add-source /path/to/nupkg
   ```
3. Run CLI: `dotnet tool run excelcli session open test.xlsx`
4. Trigger update check, verify dialog shows:
   ```
   dotnet tool update Sbroenne.ExcelMcp.CLI
   ```
   (Note: No --global flag)

## Step 10: Test Error Handling

### Test 10.1: Failed Save
1. Create session: `excelcli session open test10.xlsx`
2. Open test10.xlsx in Excel (creates file lock)
3. Right-click tray → stop daemon with "Yes" (save all)
4. Should show error dialog about save failure
5. Dialog should ask: "Stop daemon anyway?"
6. **TAKE SCREENSHOT** of error dialog
7. Click "No" to keep daemon running

### Test 10.2: Cancelled Operations
1. Verify cancel works in all dialogs:
   - Close session → Cancel keeps session open
   - Stop daemon → Cancel keeps daemon running
   - Update CLI → Cancel aborts update

## Step 11: Verify Documentation

1. Open `src/ExcelMcp.CLI/README.md`
2. Verify daemon section mentions:
   - Auto-updates feature
   - One-click update via menu
   - Save prompts with Cancel option
   - 10-minute idle timeout (not 5 minutes)
3. Open `CHANGELOG.md`
4. Verify [Unreleased] section lists all daemon improvements

## Test Results Checklist

Create a file `TEST_RESULTS.md` with the following:

```markdown
# Windows Testing Results

## Environment
- OS: Windows [version]
- .NET SDK: [version]
- Excel: [version]
- Date: [date]

## Build Status
- [ ] Build succeeded with 0 warnings
- [ ] Build succeeded with 0 errors

## Test Results

### Session Close Prompts
- [ ] Yes button saves and closes session
- [ ] No button closes without saving
- [ ] Cancel button keeps session open
- [ ] Screenshot: session_close_dialog.png

### Daemon Stop Prompts
- [ ] Yes button saves all sessions before stopping
- [ ] No button closes all without saving
- [ ] Cancel button keeps daemon running
- [ ] Stop with no sessions works without dialog
- [ ] Screenshot: daemon_stop_dialog.png

### Menu Cleanup
- [ ] No greyed-out "Excel CLI Daemon" entry visible
- [ ] Menu is clean and concise
- [ ] Screenshot: tray_menu.png

### Update Detection
- [ ] Toast notification appears when update available
- [ ] "Update to X.X.X" menu item appears
- [ ] Confirmation dialog shows correct command
- [ ] Global install shows --global flag
- [ ] Local install omits --global flag
- [ ] Screenshot: update_notification.png
- [ ] Screenshot: update_menu.png
- [ ] Screenshot: update_dialog.png

### Error Handling
- [ ] Failed save shows error dialog
- [ ] Error dialog offers continue option
- [ ] All cancel operations work correctly
- [ ] Screenshot: error_dialog.png

### Documentation
- [ ] README.md updated with new features
- [ ] CHANGELOG.md contains unreleased entry
- [ ] All documentation accurate

## Issues Found
[List any issues or bugs discovered during testing]

## Screenshots
Attach all screenshots:
1. session_close_dialog.png
2. daemon_stop_dialog.png
3. tray_menu.png
4. update_notification.png (if applicable)
5. update_menu.png (if applicable)
6. update_dialog.png (if applicable)
7. error_dialog.png

## Overall Status
- [ ] All tests passed
- [ ] Ready for merge
```

## Troubleshooting

### Issue: Daemon doesn't start
- Check if another instance is running: `tasklist | findstr excelcli`
- Kill existing: `taskkill /F /IM excelcli.exe`
- Delete lock file: `del %TEMP%\excelcli-daemon.lock`

### Issue: Tray icon not visible
- Check notification area settings (Windows may hide it)
- Look in "hidden icons" overflow menu
- Verify daemon is actually running: `excelcli daemon status`

### Issue: Dialogs don't appear
- Check if running in non-interactive session
- Verify Windows Forms is properly initialized
- Check for exceptions in Windows Event Viewer

### Issue: Update menu doesn't appear
- Verify update check completed (wait 5+ seconds after daemon start)
- Check network connectivity (NuGet API must be reachable)
- Try mocking update as described in Step 8, Option B

## Next Steps After Testing

1. Commit all screenshots to repository:
   ```powershell
   git add screenshots/*.png
   git add TEST_RESULTS.md
   git commit -m "test: add Windows testing results and screenshots"
   ```

2. Create or update the GitHub issue with:
   - Link to TEST_RESULTS.md
   - Summary of test results
   - Any issues found
   - Screenshots embedded in issue description

3. If all tests pass:
   - Mark PR as ready for review
   - Request code review from maintainers

4. If issues found:
   - Document issues clearly with steps to reproduce
   - Include error messages and screenshots
   - Suggest fixes if possible
   - Tag original developer for assistance

## Summary

This testing guide covers all six daemon improvements:
1. ✅ Copyable update instructions (menu + dialog)
2. ✅ One-click update with install detection
3. ✅ Save prompt when closing sessions
4. ✅ Removed greyed-out menu entry
5. ✅ Save prompt when stopping daemon
6. ✅ Updated documentation

Follow each step carefully and document all results with screenshots. The goal is to verify that all functionality works as designed on a real Windows system with Excel installed.

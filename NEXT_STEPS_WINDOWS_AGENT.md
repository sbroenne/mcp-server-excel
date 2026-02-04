# Issue Update: Next Steps for Windows Testing Agent

## Current Status
✅ **All code implementation is complete and committed to branch `copilot/improve-excelcli-daemon`**

All six daemon improvements have been implemented:
1. ✅ Copyable update instructions
2. ✅ One-click update with install detection  
3. ✅ Save prompt for session close
4. ✅ Removed greyed-out menu entry
5. ✅ Save prompt for daemon stop
6. ✅ Updated documentation

## Your Tasks (Windows Agent)

### Required: Build & Test on Windows

**Time Estimate:** 30-45 minutes

Since this is Windows-specific code (Windows Forms + COM), you need to:

1. **Build the project** (verify 0 warnings)
   ```powershell
   dotnet build src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release
   ```

2. **Follow the comprehensive testing guide:** 
   - See `WINDOWS_TESTING_GUIDE.md` for detailed step-by-step instructions
   - The guide covers all test scenarios with expected results
   - Each section includes checkpoints and verification steps

3. **Take screenshots** of the following dialogs:
   - Session close dialog (Yes/No/Cancel)
   - Daemon stop dialog (Yes/No/Cancel)
   - Clean tray menu (no greyed-out entry)
   - Update notification (if available)
   - Update menu item (if available)
   - Update confirmation dialog (if available)

4. **Document test results**:
   - Create `TEST_RESULTS.md` using template in testing guide
   - Mark each test as pass/fail
   - Attach all screenshots

5. **Report findings**:
   - If all tests pass → Update this issue: "✅ All tests passed, ready for merge"
   - If issues found → Document them clearly with reproduction steps

### Quick Start Commands

```powershell
# 1. Navigate to repo
cd /path/to/mcp-server-excel

# 2. Build
dotnet build src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release

# 3. Pack and install locally for testing
dotnet pack src/ExcelMcp.CLI/ExcelMcp.CLI.csproj -c Release -o ./nupkg
dotnet tool uninstall --global Sbroenne.ExcelMcp.CLI
dotnet tool install --global Sbroenne.ExcelMcp.CLI --add-source ./nupkg --version [version]

# 4. Start testing
excelcli session open test1.xlsx
excelcli session open test2.xlsx
# Now test tray icon features
```

### Critical Tests

**Must verify these scenarios:**

1. **Session Close:** Right-click tray → Sessions → file → "Close Session..."
   - Click "Yes" → saves and closes
   - Click "No" → closes without saving
   - Click "Cancel" → keeps session open

2. **Daemon Stop:** Right-click tray → "Stop Daemon" (with 2+ sessions)
   - Click "Yes" → saves all and stops
   - Click "No" → closes all and stops  
   - Click "Cancel" → keeps running

3. **Menu Check:** Right-click tray
   - Verify NO greyed-out "Excel CLI Daemon" entry
   - Menu should be clean: Sessions / [Update] / Stop

4. **Update Flow:** (if update available or mocked)
   - Toast notification mentions menu option
   - "Update to X.X.X" menu appears
   - Dialog shows correct command for install type

### Expected Build Output

```
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

### Files to Review

- Implementation: `src/ExcelMcp.CLI/Daemon/DaemonTray.cs`
- New class: `src/ExcelMcp.CLI/Infrastructure/ToolInstallationDetector.cs`
- Documentation: `src/ExcelMcp.CLI/README.md`
- Changelog: `CHANGELOG.md`

### Documentation for Reference

- `IMPLEMENTATION_SUMMARY.md` - Technical details
- `UI_MOCKUPS.md` - ASCII mockups of all dialogs
- `FINAL_SUMMARY.md` - Complete overview
- `WINDOWS_TESTING_GUIDE.md` - **START HERE** for testing steps

### Success Criteria

✅ All tests pass  
✅ All screenshots captured  
✅ No build warnings  
✅ TEST_RESULTS.md completed  
✅ PR updated with test results  

### Questions or Issues?

If you encounter any problems:
1. Check the Troubleshooting section in `WINDOWS_TESTING_GUIDE.md`
2. Document the issue with reproduction steps
3. Comment on this PR with error details
4. Include relevant error messages or screenshots

---

## Summary

**Your Action:** Run the comprehensive test suite on Windows and document results with screenshots.

**Start Here:** `WINDOWS_TESTING_GUIDE.md`

**Goal:** Verify all six daemon improvements work correctly on Windows.

**Deliverable:** `TEST_RESULTS.md` with screenshots showing all features working.

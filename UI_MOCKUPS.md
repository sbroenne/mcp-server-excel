# Excel CLI Daemon - UI Changes Mockup

## Before vs After Comparison

### Tray Menu - BEFORE
```
┌─────────────────────────────┐
│ Sessions (2)               >│
│─────────────────────────────│
│ Excel CLI Daemon     [grey] │ ← Removed (redundant)
│─────────────────────────────│
│ Stop Daemon                 │ ← Simple yes/no prompt
└─────────────────────────────┘
```

### Tray Menu - AFTER (No Update Available)
```
┌─────────────────────────────┐
│ Sessions (2)               >│
│─────────────────────────────│
│ Stop Daemon                 │ ← Now prompts to save with Cancel
└─────────────────────────────┘
```

### Tray Menu - AFTER (Update Available)
```
┌─────────────────────────────┐
│ Sessions (2)               >│
│─────────────────────────────│
│ Update to 1.6.6             │ ← NEW! One-click update
│ Stop Daemon                 │ ← Now prompts to save with Cancel
└─────────────────────────────┘
```

---

## Session Menu Changes

### Session Submenu - BEFORE
```
┌─────────────────────────────┐
│ Sessions (2)               >│
└─────────────────────────────┘
        │
        ├──> data.xlsx       >
        │    ├─ Close                  ← No prompt, just closes
        │    └─ Save & Close           ← Separate option
        │
        └──> report.xlsx    >
             ├─ Close
             └─ Save & Close
```

### Session Submenu - AFTER
```
┌─────────────────────────────┐
│ Sessions (2)               >│
└─────────────────────────────┘
        │
        ├──> data.xlsx       >
        │    └─ Close Session...      ← Opens save prompt dialog
        │
        └──> report.xlsx    >
             └─ Close Session...      ← Opens save prompt dialog
```

---

## Dialog Flows

### 1. Closing Individual Session - NEW DIALOG

**When clicking "Close Session..." on a file:**

```
┌──────────────────────────────────────────┐
│  Close Session                      [?]  │
├──────────────────────────────────────────┤
│                                          │
│  Do you want to save changes to          │
│  'data.xlsx' before closing?             │
│                                          │
│                                          │
│      [ Yes ]   [ No ]   [ Cancel ]      │
│                                          │
└──────────────────────────────────────────┘

Yes    → Save and close session
No     → Close without saving
Cancel → Keep session open (abort)
```

---

### 2. Stopping Daemon with Active Sessions - IMPROVED DIALOG

**BEFORE:**
```
┌──────────────────────────────────────────┐
│  Stop Excel CLI Daemon              [?]  │
├──────────────────────────────────────────┤
│                                          │
│  There are 2 active session(s).          │
│  Close all sessions and stop daemon?     │
│                                          │
│                                          │
│          [ Yes ]      [ No ]            │
│                                          │
└──────────────────────────────────────────┘

Yes → Closes all (no save option!)
No  → Cancel (keeps daemon running)
```

**AFTER:**
```
┌──────────────────────────────────────────┐
│  Stop Excel CLI Daemon              [?]  │
├──────────────────────────────────────────┤
│                                          │
│  There are 2 active session(s).          │
│                                          │
│  Do you want to save all sessions        │
│  before stopping the daemon?             │
│                                          │
│      [ Yes ]   [ No ]   [ Cancel ]      │
│                                          │
└──────────────────────────────────────────┘

Yes    → Save all sessions, then stop
No     → Close all without saving, then stop
Cancel → Keep daemon running (abort)
```

---

### 3. Update CLI - NEW FLOW

**Step 1: Toast Notification**
```
┌─────────────────────────────────────────────┐
│  🔔  Excel CLI Update Available             │
├─────────────────────────────────────────────┤
│                                             │
│  Version 1.6.6 is available                 │
│  (current: 1.6.5)                           │
│                                             │
│  Update via:                                │
│  dotnet tool update --global                │
│    Sbroenne.ExcelMcp.CLI                    │
│                                             │
│  Click the 'Update CLI' menu option         │
│  to update.                                 │
│                                             │
└─────────────────────────────────────────────┘

(Notification auto-dismisses after 3 seconds)
```

**Step 2: Right-click Tray Icon**
```
┌─────────────────────────────┐
│ Sessions (2)               >│
│─────────────────────────────│
│ Update to 1.6.6             │ ← NEW! Click this
│ Stop Daemon                 │
└─────────────────────────────┘
```

**Step 3: Confirmation Dialog**
```
┌──────────────────────────────────────────┐
│  Update Excel CLI                   [?]  │
├──────────────────────────────────────────┤
│                                          │
│  Update Excel CLI from 1.6.5 to 1.6.6?  │
│                                          │
│  This will run:                          │
│  dotnet tool update --global \           │
│    Sbroenne.ExcelMcp.CLI                 │
│                                          │
│  The daemon will restart after           │
│  the update.                             │
│                                          │
│          [ OK ]      [ Cancel ]         │
│                                          │
└──────────────────────────────────────────┘

OK     → Run update command
Cancel → Abort update
```

**Step 4a: Update In Progress (Balloon Tip)**
```
┌─────────────────────────────────────────────┐
│  🔔  Updating...                            │
├─────────────────────────────────────────────┤
│                                             │
│  Please wait while the CLI is updated.      │
│                                             │
└─────────────────────────────────────────────┘
```

**Step 4b: Update Success**
```
┌──────────────────────────────────────────┐
│  Update Complete                    [i]  │
├──────────────────────────────────────────┤
│                                          │
│  CLI updated successfully!               │
│                                          │
│  The daemon will now restart to use      │
│  the new version.                        │
│                                          │
│                  [ OK ]                  │
│                                          │
└──────────────────────────────────────────┘

(Daemon restarts automatically)
```

**Step 4c: Update Failed**
```
┌──────────────────────────────────────────┐
│  Update Failed                      [X]  │
├──────────────────────────────────────────┤
│                                          │
│  Update failed:                          │
│  <error message here>                    │
│                                          │
│  You can manually update by running:     │
│  dotnet tool update --global \           │
│    Sbroenne.ExcelMcp.CLI                 │
│                                          │
│                  [ OK ]                  │
│                                          │
└──────────────────────────────────────────┘
```

---

## Key Improvements Summary

### 1. ✅ Copyable Update Instructions
- **Before:** Toast notification only, hard to copy text
- **After:** Menu option + confirmation dialog shows exact command to copy

### 2. ✅ One-Click Update
- **Before:** Manual command execution required
- **After:** Click menu option → Confirm → Auto-update + restart

### 3. ✅ Session Close Prompts
- **Before:** Two separate options (Close vs Save & Close)
- **After:** Single option with 3-button dialog (Yes/No/Cancel)

### 4. ✅ Cleaner Menu
- **Before:** Redundant "Excel CLI Daemon" greyed out entry
- **After:** Clean menu without unnecessary items

### 5. ✅ Safe Daemon Stop
- **Before:** Yes/No dialog (no save option mentioned)
- **After:** Yes/No/Cancel dialog with explicit save question

### 6. ✅ Cancel Support
- **Before:** Most operations were go/no-go
- **After:** All operations support canceling mid-flow

---

## User Experience Improvements

1. **Reduced Clicks:** Update now takes 2 clicks (menu + confirm) vs typing command
2. **Clearer Intent:** All dialogs now explicitly state what will happen
3. **Data Safety:** Multiple opportunities to save before closing/stopping
4. **Cancellation:** Users can back out of any operation without consequence
5. **Transparency:** Update command is shown before execution
6. **Error Handling:** Failed updates show manual command as fallback

---

## Technical Notes

- All dialogs use standard Windows MessageBox API
- Threading: UI updates properly invoke to Windows Forms UI thread
- Installation detection: Checks executable path to determine global vs local
- Update execution: Spawns dotnet process with appropriate arguments
- Error handling: All operations have try-catch with user-friendly error messages

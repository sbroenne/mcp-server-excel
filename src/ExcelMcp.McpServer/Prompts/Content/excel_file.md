# excel_file Tool

**Actions**: open, close, create-empty, close-workbook, test

**⚠️ CRITICAL: NO 'save' ACTION**
- To persist changes, use: `action='close'` with `save=true` parameter
- Common mistake: `action='save'` (WRONG) → use `action='close', save=true` (CORRECT)

**Session lifecycle** (REQUIRED for all operations):
1. `action='open'` → returns sessionId
2. Use sessionId with other excel_* tools
3. `action='close', save=true` → persists changes and ends session
4. `action='close', save=false` → discards changes and ends session

**Related tools**:
- excel_worksheet - Add sheets after creating workbook
- excel_powerquery - Load data into workbook
- excel_range - Write data to workbook
- excel_vba - Use .xlsm extension for macro-enabled workbooks

**When to use excel_file**:
- Start/end sessions for Excel operations
- Create new blank Excel workbooks
- Validate file exists and is accessible

**Server-specific behavior**:
- open: Creates session, returns sessionId (required for all operations)
- close: Ends session, optional save parameter (default: false)
- create-empty: Creates .xlsx or .xlsm (specify extension for macro support)
- close-workbook: No-op (deprecated - automatic with single-instance architecture)
- test: Validates file without opening with Excel COM
- **File locking**: All Excel operations automatically check if file is locked and return clear error if open

**Action disambiguation**:
- open: Start session for operations
- close (save=true): Persist changes and end session
- close (save=false): Discard changes and end session
- create-empty: New blank workbook
- test: Check if file exists and has valid extension
- close-workbook: Deprecated (automatic cleanup)

**Common mistakes**:
- Using `action='save'` → WRONG! Use `action='close', save=true`
- Forgetting .xlsm for VBA → Specify extension in path
- Not starting session → All operations require sessionId from 'open' action

**Workflow optimization**:
- After create-empty: Use 'open' to start session
- After operations: Use 'close' with save=true to persist



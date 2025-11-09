# excel_file Tool

**Related tools**:
- excel_worksheet - Add sheets after creating workbook
- excel_powerquery - Load data into new workbook
- excel_range - Write initial data to new workbook
- excel_vba - Use .xlsm extension for macro-enabled workbooks

**Actions**: create-empty, close-workbook, test

**When to use excel_file**:
- Create new blank Excel workbooks
- Validate file exists and is accessible
- Use excel_worksheet after creation to add sheets
- Use excel_powerquery to populate data

**Server-specific behavior**:
- create-empty: Creates .xlsx or .xlsm (specify extension for macro support)
- close-workbook: No-op (automatic with single-instance architecture)
- test: Validates file without opening with Excel COM
- Files auto-close after each operation (no manual close needed)
- **File locking**: All Excel operations automatically check if file is locked and return clear error if open

**Action disambiguation**:
- create-empty: New blank workbook
- test: Check if file exists and has valid extension
- close-workbook: Deprecated (automatic cleanup)

**Common mistakes**:
- Manually calling close-workbook → Not needed, automatic
- Forgetting .xlsm for VBA → Specify extension in path
- Not using batch mode after creation → Start batch for multiple operations

**Error handling**:
- **File locked/open**: All operations will return clear error: "The file is already open in Excel or another process is using it. Please close the file before running automation commands."
- No need to pre-check if file is open - operations will fail gracefully with user guidance

**Workflow optimization**:
- After create-empty: Use begin_excel_batch for setup operations
- Pattern: Create file → Begin batch → Add sheets → Create queries → Commit batch



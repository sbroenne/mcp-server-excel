# excel_file Tool

**Related tools**:
- excel_worksheet - Add sheets after creating workbook
- excel_powerquery - Load data into new workbook
- excel_range - Write initial data to new workbook
- excel_vba - Use .xlsm extension for macro-enabled workbooks

**Actions**: create-empty, close-workbook, test, check-if-open

**⚠️ CRITICAL: Use check-if-open BEFORE automation!**

**When to use excel_file**:
- **check-if-open**: ALWAYS use before automation to verify file is closed
- Create new blank Excel workbooks
- Validate file exists and is accessible
- Use excel_worksheet after creation to add sheets
- Use excel_powerquery to populate data

**Server-specific behavior**:
- **check-if-open**: Proactively detects if file is locked/open (prevents automation errors)
- create-empty: Creates .xlsx or .xlsm (specify extension for macro support)
- close-workbook: No-op (automatic with single-instance architecture)
- test: Validates file without opening with Excel COM
- Files auto-close after each operation (no manual close needed)

**Action disambiguation**:
- **check-if-open**: Verify file NOT open before automation (returns isOpen: true/false with user guidance)
- create-empty: New blank workbook
- test: Check if file exists and has valid extension
- close-workbook: Deprecated (automatic cleanup)

**Common mistakes**:
- **NOT checking if file is open first** → Use check-if-open to prevent errors!
- Manually calling close-workbook → Not needed, automatic
- Forgetting .xlsm for VBA → Specify extension in path
- Not using batch mode after creation → Start batch for multiple operations

**Workflow optimization**:
- **Pre-flight**: check-if-open → if open, tell user to close → retry check
- After create-empty: Use begin_excel_batch for setup operations
- Pattern: Check if open → Create file → Begin batch → Add sheets → Create queries → Commit batch


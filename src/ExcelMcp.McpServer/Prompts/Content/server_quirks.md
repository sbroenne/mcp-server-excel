# Server Quirks & Gotchas - Things I Need to Remember

## CRITICAL: File Access Requirements

**NEVER work on files that are already open in Excel**
- ExcelMcp requires EXCLUSIVE access to workbooks
- If file is open in Excel UI or another process → operations WILL FAIL
- Error: "The file is already open in Excel or another process is using it"
- **USER ACTION REQUIRED**: Tell user to close the file first!
- This is NOT optional - Excel COM automation requires exclusive access

**Why this matters:**
- Excel uses file locking to prevent corruption
- Multiple processes accessing same file = data loss risk
- Automation needs predictable state (no user edits during operation)

**How to detect:**
- User says "the file is open" or "I have Excel running"
- Error message mentions "already open" or "locked by another process"
- ALWAYS tell user: "Please close the file in Excel before running automation"

## Data Type Surprises

**Single cells return 2D arrays, not scalars**
- excel_range(get-values, rangeAddress='A1') → [[42]] not 42
- ALWAYS expect [[value]] even for single cell
- I must extract value[0][0] if I need the scalar

**Named ranges use empty sheetName**
- excel_range(rangeAddress='SalesData', sheetName='') ← empty string!
- NOT sheetName='Sheet1' for named ranges

## Excel COM Limitations

**Cannot create OLEDB/ODBC connections programmatically**
- excel_connection can only MANAGE existing connections
- User must create OLEDB/ODBC in Excel UI first
- TEXT connections work fine for automation

**Cannot delete last worksheet**
- Excel always requires at least one sheet
- excel_worksheet(delete) will fail if it's the last one

**VBA requires .xlsm files**
- excel_vba won't work on .xlsx
- File must be macro-enabled

## Load Destination Confusion

**'worksheet' vs 'data-model' vs 'both'**
- worksheet: Users see data, NO DAX capability
- data-model: Ready for DAX, users DON'T see data
- both: Users see data AND DAX works
- DEFAULT is 'worksheet' if not specified

**Cannot directly add worksheet query to Data Model**
- If loaded to worksheet only, can't use excel_table add-to-datamodel
- Must use excel_powerquery set-load-to-data-model to fix

## Number Format Edge Cases

**Format codes are strings, not patterns**
- '$#,##0.00' is exact string, not regex
- Common codes: see format_codes.md completion

**set-number-format vs set-number-formats (plural)**
- set-number-format: ONE format for entire range
- set-number-formats: DIFFERENT format per cell (2D array)

## Common Error Patterns

**"Value does not fall within expected range"**
- Usually: Trying to create OLEDB connection (can't do it)
- Or: Invalid range address
- Or: Operation not supported by Excel COM

**"Refresh failed or data not updated"**
- RefreshAll() is async and unreliable
- Solution: Use individual query/connection refresh (synchronous)

**"Parameter name missing ="**
- Named ranges must be: =Sheet1!$A$1
- NOT: Sheet1!$A$1 (missing = prefix)

# Server Quirks & Gotchas - Things I Need to Remember

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

## Batch Mode Imperatives

**Batch sessions MUST be committed**
- begin_excel_batch creates resource that MUST be released
- Forgetting commit_excel_batch = resource leak
- Always pair begin with commit

**One batch per file**
- Cannot create multiple batches for same file
- Must commit first batch before starting new one

**Batch mode is UPFRONT decision**
- Must call begin_excel_batch BEFORE first operation
- Cannot "upgrade" to batch mode mid-workflow
- Keyword detection should happen immediately

## Number Format Edge Cases

**Format codes are strings, not patterns**
- '$#,##0.00' is exact string, not regex
- Common codes: see format_codes.md completion

**set-number-format vs set-number-formats (plural)**
- set-number-format: ONE format for entire range
- set-number-formats: DIFFERENT format per cell (2D array)

## Refresh vs LoadDestination

**refresh action can apply loadDestination**
- Old way: set-load-to-table + refresh (2 calls)
- New way: refresh(loadDestination='worksheet') (1 call)
- Saves time for connection-only queries

## Common Error Patterns

**"Value does not fall within expected range"**
- Usually: Trying to create OLEDB connection (can't do it)
- Or: Invalid range address
- Or: Operation not supported by Excel COM

**"QueryTable not found after refresh"**
- Using RefreshAll() instead of queryTable.Refresh(false)
- Solution: Server uses synchronous refresh (already handled)

**"Parameter name missing ="**
- Named ranges must be: =Sheet1!$A$1
- NOT: Sheet1!$A$1 (missing = prefix)

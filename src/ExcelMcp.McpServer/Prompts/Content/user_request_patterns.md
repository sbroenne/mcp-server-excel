# Common User Request Patterns - How to Interpret

## 🚨 PRE-FLIGHT: File Access Check

**BEFORE processing ANY request, check if file is accessible:**

**User signals file might be open:**
- "I have the file open" → STOP, tell them to close it
- "The file is currently open in Excel" → STOP
- "Can I run this while viewing the file?" → NO, tell them to close first
- "Error says file is locked" → File is open, tell them to close it

**Always ask if unsure:**
- "Is the Excel file currently open?"
- "Please close the file before we proceed with automation"

**This is mandatory - Excel COM automation requires exclusive file access!**

## Data Import Requests

**"Load this CSV file"**
→ excel_powerquery(create, sourcePath='file.csv', loadDestination='worksheet')
→ excel_querytable(create-from-connection) for simple imports (no M code)
→ NOT excel_table (that's for existing data)

**"Import data from SQL Server"**
→ excel_connection(create) with OLEDB connection string
→ Then excel_querytable(create-from-connection) for simple import
→ Or excel_powerquery for complex transformations

**"Load Power Query results to worksheet"**
→ excel_querytable(create-from-query) - simpler than excel_powerquery load-to
→ OR excel_powerquery(create, loadDestination='worksheet')

**"Put data in Data Model for DAX"**
→ excel_powerquery(create, loadDestination='data-model')
→ NOT 'worksheet' (that won't work for DAX)

**"Refresh data from external source"**
→ excel_querytable(refresh) - synchronous, guaranteed persistence
→ excel_querytable(refresh-all) - all QueryTables in workbook

## Formatting Requests

**"Make headers bold with blue background"**
→ excel_range(format-range, bold=true, fillColor='#4472C4', fontColor='#FFFFFF')
→ Single call with multiple properties

**"Format column D as currency"**
→ excel_range(set-number-format, rangeAddress='D:D', formatCode='$#,##0.00')

**"Add dropdown for Status column"**
→ excel_range(validate-range, validationType='list', validationFormula1='Active,Inactive,Pending')

## Analytics Requests

**"Create Total Sales measure"**
→ First: Check data in Data Model with excel_datamodel(list-tables)
→ Then: excel_datamodel(create-measure, daxFormula='SUM(Sales[Amount])')

**"Link Sales to Products table"**
→ excel_datamodel(create-relationship, fromTable='Sales', fromColumn='ProductID', toTable='Products', toColumn='ProductID')

## Configuration Requests

**"Set up parameters for date range"**
→ excel_namedrange(create-bulk) with StartDate and EndDate
→ NOT excel_range (parameters are named ranges)

## Structure Requests

**"Create new sheet called Reports"**
→ excel_worksheet(create, sheetName='Reports')

**"Convert this data to a table"**
→ excel_table(create, sourceRange='A1:E100', tableName='SalesData')

## VBA Requests

**"Export VBA code for version control"**
→ excel_vba(export, moduleName='Module1', targetPath='Module1.bas')

**"Import macro from file"**
→ excel_vba(import, sourcePath='Module1.bas')
→ File must be .xlsm

## Discovery Requests

**"What Power Queries are in this file?"**
→ excel_powerquery(list)

**"What QueryTables exist?"**
→ excel_querytable(list)

**"Show me all DAX measures"**
→ excel_datamodel(list-measures)

**"What sheets exist?"**
→ excel_worksheet(list)

**"What connections are available?"**
→ excel_connection(list)

**"Are there any active batch sessions?"**
→ list_excel_batches

## Edge Case Interpretations

**"Delete all data"**
→ excel_range(clear-contents) NOT clear-all (preserve formatting)

**"Get data from A1"**
→ Remember: Returns [[value]] not value
→ Extract with result.values[0][0] if needed

**"Hide this sheet from users"**
→ excel_worksheet(very-hide) for strong protection
→ excel_worksheet(hide) for normal hiding

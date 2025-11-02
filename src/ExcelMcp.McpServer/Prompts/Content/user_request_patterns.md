# Common User Request Patterns - How to Interpret

## Data Import Requests

**"Load this CSV file"**
→ excel_powerquery(import, sourcePath='file.csv', loadDestination='worksheet')
→ NOT excel_table (that's for existing data)

**"Import data from SQL Server"**
→ User must create connection in Excel UI first (OLEDB limitation)
→ Then excel_connection(refresh) or excel_powerquery

**"Put data in Data Model for DAX"**
→ excel_powerquery(import, loadDestination='data-model')
→ NOT 'worksheet' (that won't work for DAX)

## Bulk Operation Requests

**"Import these 4 files"** (number = batch!)
→ begin_excel_batch
→ excel_powerquery × 4 with batchId
→ commit_excel_batch

**"Create measures for Sales, Revenue, Profit"** (list = batch!)
→ begin_excel_batch
→ excel_datamodel(create-measure) × 3 with batchId
→ commit_excel_batch

**"Add parameters: StartDate, EndDate, Region"** (list = batch!)
→ excel_namedrange(create-bulk) with JSON array (no batch needed, already batched)

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

**"Show me all DAX measures"**
→ excel_datamodel(list-measures)

**"What sheets exist?"**
→ excel_worksheet(list)

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

# BEFORE CREATING EXCEL TABLE - GATHER THIS INFO

REQUIRED:
☐ Excel file path
☐ Worksheet name
☐ Source range address (existing data to convert, e.g., 'A1:E100')
☐ Table name (unique identifier)

RECOMMENDED:
☐ Has header row? (true/false, default true)
☐ Table style (e.g., 'TableStyleMedium2', 'TableStyleLight1')
☐ Show totals row? (true/false, default false)

OPTIONAL:
☐ Add to Data Model? (use excel_table add-to-datamodel action after creation)
☐ Initial filters? (apply after creation with apply-filter action)
☐ Sort columns? (apply after creation with sort action)

WORKFLOW OPTIMIZATION:
☐ Creating multiple tables? → Use batch mode (begin_excel_batch)
☐ Pattern: Create table → Apply filters → Sort → Format columns → Add to data model

PREREQUISITES:
☐ Data already exists in worksheet range (use excel_range if data needs to be added first)
☐ Range has consistent columns (table requires structured data)

ASK USER FOR MISSING INFO before calling excel_table(action: 'create')

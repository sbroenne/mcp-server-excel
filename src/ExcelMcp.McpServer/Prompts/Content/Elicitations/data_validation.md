# BEFORE ADDING DATA VALIDATION - GATHER THIS INFO

REQUIRED:
☐ Excel file path
☐ Worksheet name
☐ Range address (cells to validate)
☐ Validation type:
  - 'list' (dropdown with fixed choices)
  - 'decimal' (number with min/max)
  - 'whole' (integer with min/max)
  - 'date' (date range)
  - 'time' (time range)
  - 'textLength' (character count limits)
  - 'custom' (formula-based validation)

TYPE-SPECIFIC INFO:

FOR LIST VALIDATION (creates dropdown):
☐ List values (e.g., ['Active', 'Inactive', 'Pending'])
☐ Show dropdown? (true recommended)
⚠️ WORKFLOW: 
  1. First write values to a worksheet range (e.g., Sheet1!$Z$1:$Z$10)
  2. Then set formula1 to reference that range (e.g., '=$Z$1:$Z$10')
  3. This creates a proper dropdown that users can select from
☐ Note: Comma-separated strings do NOT create dropdowns - only range references do!

FOR NUMBER VALIDATION (decimal/whole):
☐ Operator (between, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, equal, notEqual)
☐ Minimum value
☐ Maximum value (if 'between' operator)

FOR DATE/TIME VALIDATION:
☐ Operator (same as number)
☐ Start date/time
☐ End date/time (if 'between')

FOR TEXT LENGTH VALIDATION:
☐ Operator (same as number)
☐ Character count limit

FOR CUSTOM VALIDATION:
☐ Excel formula returning TRUE/FALSE (e.g., '=MOD(A1,5)=0')

OPTIONAL (RECOMMENDED):
☐ Error style (stop, warning, information)
☐ Error title and message (help users fix mistakes)
☐ Input message (show when cell selected)

ASK USER FOR MISSING INFO before calling excel_range(action: 'validate-range')

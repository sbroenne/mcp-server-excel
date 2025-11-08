# excel_table - Server Quirks

**Action disambiguation**:
- add-to-datamodel vs loadDestination: add-to-datamodel for existing Excel tables, loadDestination for Power Query imports
- resize: Changes table boundaries (not the same as append)
- get-structured-reference: Returns formula + range address for use with excel_range

**Server-specific quirks**:
- Column names: Any string including purely numeric (e.g., "60")
- Table names: Must start with letter/underscore, alphanumeric only
- AutoFilter: Enabled by default on table creation
- Structured references: =Table1[@ColumnName] (auto-adjusts when table resizes)

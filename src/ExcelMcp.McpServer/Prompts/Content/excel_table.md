# excel_table - Server Quirks

**Action disambiguation**:
- add-to-datamodel vs loadDestination: add-to-datamodel for existing Excel tables, loadDestination for Power Query imports
- resize: Changes table boundaries (not the same as append)
- get-structured-reference: Returns formula + range address for use with excel_range
- set-style: Apply built-in Excel table styles (60+ presets available)

**Server-specific quirks**:
- Column names: Any string including purely numeric (e.g., "60")
- Table names: Must start with letter/underscore, alphanumeric only
- AutoFilter: Enabled by default on table creation
- Structured references: =Table1[@ColumnName] (auto-adjusts when table resizes)
- Table styles: Can be applied during create or changed later with set-style

**Table Style Categories** (60+ built-in styles):
- **Light** (TableStyleLight1-21): Subtle colors, minimal formatting
- **Medium** (TableStyleMedium1-28): Balanced colors, most popular
- **Dark** (TableStyleDark1-11): High contrast, bold colors

**Popular table styles**:
- TableStyleMedium2 - orange accents (most common)
- TableStyleMedium9 - gray/blue professional
- TableStyleLight9 - blue banding (subtle)
- TableStyleDark1 - dark blue header (high contrast)

**Style behavior**:
- Styles include header formatting, row banding, total row formatting
- Changing style preserves data and structure
- Empty string "" removes table styling

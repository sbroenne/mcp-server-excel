# 9-Command Architecture Plan

## Overview
Consolidate ExcelMcp from ~94 CLI commands into 9 logical command groups aligned with Excel's actual object model.

## Command Structure

### 1. `range` - Cell/Range Operations
**Rationale:** In Excel, cells and ranges are the same object type. A1 is a 1x1 range. Formatting, validation, and conditional formatting all apply to ranges.

**Actions:**
- `read` - Read cell/range values
- `write` - Write cell/range values  
- `get-formula` - Get cell/range formulas
- `set-formula` - Set cell/range formulas
- `format-background` - Set background color
- `format-font` - Set font properties
- `format-border` - Apply borders
- `format-number` - Set number format
- `format-alignment` - Set alignment
- `clear-format` - Remove formatting
- `add-validation` - Add data validation rule
- `list-validation` - List validation rules
- `delete-validation` - Remove validation
- `add-conditional-format` - Add conditional formatting rule
- `list-conditional-formats` - List conditional formatting rules
- `delete-conditional-format` - Remove conditional formatting
- `add-hyperlink` - Add hyperlink to cell/range
- `delete-hyperlink` - Remove hyperlink
- `add-comment` - Add comment/note
- `delete-comment` - Remove comment/note

**Consolidates:** cell, cell-set-background-color, cell-set-font-color, cell-set-font, cell-set-border, cell-set-number-format, cell-set-alignment, cell-clear-formatting, cell-get-value, cell-set-value, cell-get-formula, cell-set-formula

### 2. `table` - Excel Tables (ListObjects)
**Rationale:** Excel Tables are distinct structured objects with unique lifecycle and operations.

**Actions:**
- `create` - Create table from range
- `list` - List all tables
- `info` - Get table metadata
- `rename` - Rename table
- `delete` - Convert table back to range
- `resize` - Expand/shrink table range
- `toggle-totals` - Enable/disable totals row
- `set-column-total` - Configure column aggregation
- `read` - Read table data
- `append` - Add rows to table
- `set-style` - Change table style
- `add-to-datamodel` - Add to Power Pivot

**Keeps:** All existing table commands as-is

### 3. `sheet` - Worksheet Operations
**Rationale:** Worksheets are containers for ranges/tables. Protection is sheet-level.

**Actions:**
- `list` - List all worksheets
- `create` - Create worksheet
- `rename` - Rename worksheet
- `copy` - Copy worksheet
- `delete` - Delete worksheet
- `clear` - Clear worksheet content
- `read` - Read worksheet range
- `write` - Write to worksheet range
- `append` - Append rows to worksheet
- `protect` - Protect worksheet with password/permissions
- `unprotect` - Remove worksheet protection
- `get-protection` - Query protection status

**Consolidates:** All sheet-* commands + protection operations

### 4. `powerquery` - Power Query M Code
**Rationale:** Power Query is a distinct feature with its own lifecycle.

**Actions:**
- `list` - List queries
- `view` - View M code
- `import` - Import query from file
- `export` - Export query to file
- `update` - Update query M code
- `delete` - Delete query
- `refresh` - Refresh query
- `set-load-config` - Configure load destination

**Keeps:** All existing powerquery commands as-is

### 5. `file` - Workbook File Operations
**Rationale:** File-level operations.

**Actions:**
- `create-empty` - Create new empty workbook
- `close-workbook` - Close workbook from pool

**Keeps:** All existing file commands as-is

### 6. `parameter` - Named Ranges
**Rationale:** Named ranges serve as parameters/configuration.

**Actions:**
- `list` - List named ranges
- `get` - Get named range value
- `set` - Set named range value
- `create` - Create named range
- `delete` - Delete named range

**Keeps:** All existing parameter commands as-is

### 7. `vba` - VBA Macros
**Rationale:** VBA is a distinct programming layer.

**Actions:**
- `list` - List VBA modules
- `export` - Export VBA to file
- `import` - Import VBA from file
- `update` - Update VBA code
- `run` - Execute VBA macro
- `delete` - Delete VBA module

**Keeps:** All existing vba commands as-is

### 8. `chart` - Charts and Graphs
**Rationale:** Charts are visual objects with distinct lifecycle.

**Actions:**
- `create` - Create chart
- `list` - List charts
- `delete` - Delete chart
- `update-data` - Update chart data range
- `set-type` - Change chart type
- `set-style` - Apply chart style
- `set-title` - Set chart title
- `set-axis` - Configure axis
- `set-legend` - Configure legend

**Status:** NEW - To be implemented

### 9. `pivot` - PivotTables
**Rationale:** PivotTables are complex analytical objects.

**Actions:**
- `create` - Create PivotTable
- `list` - List PivotTables
- `refresh` - Refresh PivotTable data
- `delete` - Delete PivotTable
- `add-field` - Add field to PivotTable
- `set-filter` - Set PivotTable filter
- `set-aggregation` - Configure field aggregation

**Status:** NEW - To be implemented

## Migration Status

### Completed
- ✅ Tool naming (excel_ prefix removed)
- ✅ Table commands (full lifecycle)
- ✅ Cell formatting (Core layer)
- ✅ Sheet protection (Core layer)

### In Progress  
- 🔄 Range consolidation (merge cell commands)
- 🔄 Sheet consolidation (merge protection)

### Pending
- ⏳ Hyperlinks (add to range)
- ⏳ Comments (add to range)
- ⏳ Data Validation (add to range)
- ⏳ Conditional Formatting (add to range)
- ⏳ Charts (new command)
- ⏳ PivotTables (new command)
- ⏳ Formulas/Calculations (enhance range)

## Benefits

1. **Reduced Command Count:** 94 → 9 commands (90% reduction)
2. **Aligned with Excel:** Matches Excel's actual object model
3. **LLM-Friendly:** Fits within typical LLM tool limits (usually 10-20 tools)
4. **Intuitive:** Operations grouped by target object type
5. **Extensible:** Easy to add new actions to existing commands

## Implementation Plan

1. **Phase 1:** Merge cell → range (consolidate formatting operations)
2. **Phase 2:** Merge protection → sheet
3. **Phase 3:** Add missing range operations (hyperlinks, comments, validation, conditional formatting)
4. **Phase 4:** Implement chart command
5. **Phase 5:** Implement pivot command  
6. **Phase 6:** Update all tests and documentation

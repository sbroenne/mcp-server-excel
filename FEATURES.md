# ExcelMcp - Complete Feature Reference

**12 specialized tools with 172 operations for comprehensive Excel automation**

---

## 🔄 Power Query & M Code (9 operations)

**✨ Atomic Operations** - Single-call workflows replace multi-step patterns:
- **Create** - Import + load in one operation (replaces import → load workflow)
- **Update & Refresh** - Update M code + refresh data atomically
- **Refresh All** - Batch refresh all queries in workbook
- **Update M Code** - Stage code changes without refreshing data
- **Unload** - Convert loaded query to connection-only

**Core Operations:**
- List, view, delete Power Query transformations
- Manage query load destinations (worksheet/data model/connection-only/both)
- Get load configuration for existing queries
- List Excel workbook sources for Power Query integration

---

## 📊 Data Model & DAX (Power Pivot) (14 operations)

- Create/update/delete DAX measures with format types (Currency, Percentage, Decimal, General)
- Manage table relationships (create, toggle active/inactive, delete)
- Discover model structure (tables, columns, measures, relationships)
- Get comprehensive model information
- Refresh data model
- **Note:** DAX calculated columns are not supported (use Excel UI for calculated columns)

---

## 🎨 Excel Tables (ListObjects) (24 operations)

- **Lifecycle:** create, resize, rename, delete, get info
- **Styling:** apply table styles, toggle totals row, set column totals
- **Column Management:** add, remove, rename columns
- **Data Operations:** append rows, apply filters (criteria/values), clear filters, get filter state
- **Sorting:** single-column sort, multi-column sort (up to 3 levels)
- **Number Formatting:** get/set column number formats
- **Advanced Features:** structured references, Data Model integration, get table data with optional visible-only filtering

---

## 📈 PivotTables (25 operations)

- **Creation:** create from ranges, Excel Tables, or Data Model
- **Field Management:** add/remove fields to Row, Column, Value, Filter areas
- **Aggregation Functions:** Sum, Average, Count, Min, Max, etc. with validation
- **Advanced Features:** field filters, sorting, custom field names, number formatting
- **Data Extraction:** get PivotTable data as 2D arrays for further analysis
- **Lifecycle:** list, get info, delete, refresh

---

## 📉 Charts (14 operations)

- Create charts from ranges, tables, or PivotTables
- Move charts across worksheets or to dedicated chart sheets
- Manage series (add, update, delete) with category name/values control
- Configure data sources, chart types, and set dynamic ranges
- Update chart titles, axes, legend, plot area, and formatting
- List charts, retrieve chart info, and delete when no longer needed

---

## 📝 VBA Macros (6 operations)

- List all VBA modules and procedures
- View module code without exporting
- Export/import VBA modules to/from files
- Update existing modules
- Execute macros with parameters
- Delete modules
- Version control VBA code through file exports

---

## 📋 Ranges (42 operations)

### Data Operations (10 actions)
- Get/set values and formulas
- Clear (all/contents/formats)
- Copy/paste (all/values/formulas)
- Insert/delete rows/columns/cells
- Find/replace
- Sort

### Number Formatting (3 actions)
- Get formats as 2D arrays
- Apply uniform format
- Set individual cell formats

### Visual Formatting (1 action)
- Font (name, size, bold, italic, underline, color)
- Fill color
- Borders (style, weight, color)
- Alignment (horizontal, vertical, wrap text, orientation)

### Data Validation (3 actions)
- Add validation rules (dropdowns, number/date/text rules)
- Get validation info
- Remove validation

### Hyperlinks (4 actions)
- Add, remove, list all, get specific hyperlink

### Smart Range Operations (3 actions)
- UsedRange, CurrentRegion, get range info (address, dimensions, format)

### Merge Operations (3 actions)
- Merge cells, unmerge cells, get merge info

### Auto-Sizing (2 actions)
- Auto-fit columns, auto-fit rows

### Conditional Formatting (2 actions)
- Add conditional formatting rules
- Clear conditional formatting

### Cell Protection (2 actions)
- Set cell lock status
- Get cell lock status

### Formatting & Styling (3 actions)
- Get style, set style, format range

---

## 📄 Worksheets (16 operations)

- **Lifecycle:** create, rename, copy, move, delete
- **Copy/Move Between Workbooks:** cross-workbook operations
- **Tab Colors:** set, get, clear (RGB values)
- **Visibility Controls:** show, hide, very-hide, get/set status

---

## 🔌 Data Connections (9 operations)

- Create and manage OLEDB/ODBC connections (when provider is installed)
- Refresh connections, test connectivity, and control refresh settings
- Load connection-only sources to worksheet tables when supported
- Update connection strings, command text, and metadata
- Automatically redirect TEXT/WEB scenarios to `excel_powerquery` for reliable imports

**Note:** OLEDB/ODBC creation requires the provider to be installed (e.g., Microsoft.ACE.OLEDB.16.0, SQLOLEDB)

---

## 🏷️ Named Ranges (6 operations)

- List all named ranges with references
- Get or set single values (ideal for parameter automation)
- Create, update, or delete named ranges individually or in bulk
- Maintain workbook parameters without touching worksheets

---

## 📁 File Operations (5 operations)

- **Session Management:** open, close (with optional save)
- **Create Empty:** new .xlsx or .xlsm workbooks
- **Test:** verify workbook can be opened
- **💡 Show Excel Mode:** Open with `showExcel:true` to watch AI changes live - perfect for debugging, demos, and learning

---

## 🎨 Conditional Formatting (2 operations)

- **Add Rules:** cell value comparisons, expression-based formulas
- **Clear Rules:** remove formatting from ranges

---

## 📊 Total Operations Summary

| Category | Operations |
|----------|-----------|
| Power Query | 9 |
| Data Model/DAX | 14 |
| Excel Tables | 24 |
| PivotTables | 25 |
| Charts | 14 |
| VBA Macros | 6 |
| Ranges | 42 |
| Worksheets | 16 |
| Connections | 9 |
| Named Ranges | 6 |
| File Operations | 5 |
| Conditional Formatting | 2 |
| **Total** | **172** |

---

## 🚀 Growing Feature Set

ExcelMcp is actively developed with new features added regularly. Check the [releases page](https://github.com/sbroenne/mcp-server-excel/releases) for the latest additions.

**Recent Additions:**
- Chart automation (14 operations)
- Atomic Power Query operations (create, update-and-refresh)
- PivotTable data extraction
- Conditional formatting rules
- Cross-workbook worksheet operations

---

## 📚 Documentation

- **[Installation Guide](docs/INSTALLATION.md)** - Setup for all AI assistants
- **[MCP Server Guide](src/ExcelMcp.McpServer/README.md)** - Tool documentation and examples
- **[CLI Guide](src/ExcelMcp.CLI/README.md)** - Command-line reference
- **[Contributing](docs/CONTRIBUTING.md)** - Development guidelines

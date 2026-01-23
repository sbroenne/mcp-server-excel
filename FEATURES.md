# ExcelMcp - Complete Feature Reference

**22 specialized tools with 206 operations for comprehensive Excel automation**

---

## 📁 File Operations (6 operations)

- **List Sessions:** View all active Excel sessions
- **Open:** Open workbook and create session (returns session ID for all subsequent operations)
- **Close:** Close session with optional save
- **Close Workbook:** Close workbook without closing Excel
- **Create Empty:** Create new .xlsx or .xlsm workbook
- **Test:** Verify workbook can be opened and is accessible

---

## 🔄 Power Query & M Code (10 operations)

**Atomic Operations** - Single-call workflows:
- **List:** List all Power Query queries in workbook
- **View:** View the M code of a Power Query
- **Create:** Import + load in one operation (atomic workflow) with automatic formatting
- **Update:** Update M code with automatic formatting and auto-refresh
- **Rename:** Rename a Power Query (trim + case-insensitive uniqueness check)
- **Refresh:** Refresh a Power Query with timeout detection
- **Refresh All:** Batch refresh all queries in workbook
- **Load To:** Configure load destination and refresh (atomic)
- **Get Load Config:** Get current load configuration
- **Delete:** Remove Power Query from workbook

**Automatic M-Code Formatting:** M code is automatically formatted on write operations (Create, Update) using the powerqueryformatter.com API (by mogularGmbH, MIT License). Read operations return M code as stored in Excel. Formatting adds ~100-500ms network latency but dramatically improves readability with proper indentation, spacing, and line breaks. Graceful fallback returns original M code if formatting fails.

---

## 📊 Data Model & DAX (Power Pivot) (19 operations)

- **List Tables:** Discover all tables in the Data Model
- **Read Table:** Get specific table information
- **Rename Table:** Rename a Data Model table (best-effort via Power Query; returns clear error if not supported)
- **List Columns:** List columns for a table
- **List Measures:** List all DAX measures with formula previews
- **Read Info:** Get comprehensive model information
- **Create Measure:** Create new DAX measure with automatic formatting (format types: Currency, Percentage, Decimal, General)
- **Update Measure:** Modify existing measure with automatic formatting
- **Delete Measure:** Remove measure from model
- **Delete Table:** Remove table from Data Model
- **List Relationships:** View all table relationships
- **Read Relationship:** Get specific relationship info
- **Create Relationship:** Create relationship between tables
- **Update Relationship:** Modify relationship (toggle active/inactive)
- **Delete Relationship:** Remove relationship
- **Refresh:** Refresh entire Data Model
- **List Workbook Connections:** List Power Query sources available for integration
- **Evaluate:** Execute DAX EVALUATE queries and return tabular results (for ad-hoc analysis)
- **Execute DMV:** Execute SQL-like DMV (Dynamic Management View) queries for metadata discovery

**Automatic DAX Formatting:** DAX formulas are automatically formatted on write operations (CreateMeasure, UpdateMeasure) using the official Dax.Formatter library (SQLBI). Read operations return raw DAX as stored in Excel. Formatting adds ~100-500ms network latency but dramatically improves readability. Graceful fallback returns original DAX if formatting fails.

**Note:** DAX calculated columns not supported - use Excel UI for calculated columns

---

## 🎨 Excel Tables (ListObjects) (27 operations)

**Lifecycle:**
- List, read, create, rename, resize, delete tables

**Styling & Formatting:**
- Apply table styles
- Toggle totals row
- Set column totals

**Data Operations:**
- Append rows
- Get table data (with optional visible-only filtering)
- Add to Data Model

**DAX-Backed Tables:**
- Create from DAX (create Excel Table populated by DAX EVALUATE query)
- Update DAX (change the DAX query of an existing DAX-backed table)
- Get DAX (retrieve DAX query info from a table)

**Filter Operations:**
- Apply filter (criteria)
- Apply filter (values)
- Clear filters
- Get filter state

**Column Management:**
- Add, remove, rename columns

**Structured References:**
- Get structured reference (formula syntax for table columns/ranges)

**Sorting:**
- Single-column sort
- Multi-column sort (up to 3 levels)

**Number Formatting:**
- Get column number formats
- Set column number formats

---

## 📈 PivotTables (30 operations)

**Creation:**
- Create from range
- Create from Excel Table
- Create from Data Model

**Field Management:**
- List all fields (row, column, value, filter areas)
- Add row field, column field, value field, filter field
- Remove field

**Field Configuration:**
- Set field aggregation function (Sum, Average, Count, Min, Max, etc.)
- Set custom field name
- Set field number format
- Set field filter criteria
- Sort field (ascending/descending)

**Calculated Fields (Regular PivotTables):**
- List calculated fields
- Create calculated field
- Delete calculated field

**Calculated Members (OLAP/Data Model PivotTables):**
- List calculated members
- Create calculated member
- Delete calculated member

**Layout & Formatting:**
- Set layout (table or outline)
- Set subtotals display
- Set grand totals display

**Data Operations:**
- Get PivotTable data as 2D array
- Refresh PivotTable

**Lifecycle:**
- List PivotTables
- Read PivotTable info
- Delete PivotTable

---

## 📉 Charts (26 operations)

**Creation:**
- Create from range
- Create from PivotTable

**Series Management:**
- Add series
- Remove series
- Update series data

**Configuration:**
- Set data source range
- Set chart type
- Show/hide legend
- Set style

**Formatting:**
- Set chart title
- Set axis title
- Set axis number format
- Get axis number format

**Data Labels:**
- Configure data labels (show values, percentages, category names, etc.)
- Set label position (Center, InsideEnd, OutsideEnd, etc.)
- Apply to all series or specific series

**Axis Scale:**
- Get axis scale settings
- Set minimum/maximum scale
- Set major/minor units

**Gridlines:**
- Get gridlines configuration
- Set major/minor gridlines visibility

**Series Formatting:**
- Set marker style (Circle, Square, Diamond, Triangle, etc.)
- Set marker size
- Set marker colors

**Trendlines:**
- Add trendline (Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage)
- List trendlines on series
- Delete trendline
- Configure trendline (forecast forward/backward, display equation, display R²)

**Lifecycle:**
- List charts
- Read chart info
- Move chart (to different worksheet or new sheet)
- Delete chart

---

## 📋 Ranges (42 operations)

**Data Operations:**
- Get values
- Set values
- Get formulas
- Set formulas
- Clear all
- Clear contents
- Clear formats
- Copy
- Copy values
- Copy formulas
- Insert cells
- Delete cells
- Insert rows
- Delete rows
- Insert columns
- Delete columns
- Find
- Replace
- Sort

**Discovery & Utilities:**
- Get used range
- Get current region
- Get range info (address, dimensions)

**Hyperlinks:**
- Add hyperlink
- Remove hyperlink
- List hyperlinks
- Get specific hyperlink

**Number Formatting:**
- Get number formats (as 2D array)
- Set number format (uniform)
- Set number formats (individual)

**Visual Formatting:**
- Get style
- Set style (built-in Excel styles)
- Format range (font, color, borders, alignment, orientation)

**Data Validation:**
- Add validation rules (dropdowns, number/date/text rules)
- Get validation info
- Remove validation

**Merge Operations:**
- Merge cells
- Unmerge cells
- Get merge info

**Cell Protection:**
- Set cell lock status
- Get cell lock status

**Auto-Sizing:**
- Auto-fit columns
- Auto-fit rows

---

## 📄 Worksheets (16 operations)

**Lifecycle:**
- List worksheets
- Create worksheet
- Rename worksheet
- Copy worksheet
- Move worksheet
- Delete worksheet

**Cross-Workbook Operations:**
- Copy worksheet to file (atomic)
- Move worksheet to file (atomic)

**Tab Colors:**
- Set tab color (RGB)
- Get tab color
- Clear tab color

**Visibility:**
- Show worksheet
- Hide worksheet
- Very hide worksheet (hidden from UI)
- Get visibility status
- Set visibility status

---

## 🔌 Data Connections (9 operations)

- **List:** View all data connections
- **View:** Get connection details
- **Create:** Create OLEDB/ODBC connections (requires provider installed)
- **Test:** Verify connection validity
- **Refresh:** Refresh connection data
- **Delete:** Remove connection
- **Load To:** Load connection data to worksheet (when supported)
- **Get Properties:** Get connection string and metadata
- **Set Properties:** Update connection string, command text, and settings

**Supported Types:**
- OLEDB (requires Microsoft.ACE.OLEDB.16.0 or similar)
- ODBC (requires ODBC driver installed)
- Power Query connections (atomic redirect to excel_powerquery)

**Automatic Fallback:**
- TEXT/WEB connections automatically redirect to excel_powerquery for reliable imports

---

## 🏷️ Named Ranges (Parameters) (6 operations)

- **List:** List all named ranges with references
- **Read:** Get value of a named range
- **Write:** Set value of a named range (ideal for parameter automation)
- **Create:** Create new named range
- **Update:** Modify existing named range
- **Delete:** Remove named range

**Use Cases:**
- Workbook parameter management without touching worksheets
- Ideal for automation: update parameter → Power Query refreshes automatically

---

## 📝 VBA Macros (6 operations)

- **List:** List all VBA modules and procedures
- **View:** Display module code without exporting
- **Import:** Add VBA module from file
- **Update:** Modify existing VBA module
- **Delete:** Remove VBA module
- **Run:** Execute macro with optional parameters

**Features:**
- Version control through file exports
- Parameter passing to macros
- Full module lifecycle management

---

## �️ Slicers (8 operations)

**PivotTable Slicers:**
- **Create Slicer:** Add slicer for PivotTable field with optional position
- **List Slicers:** List all PivotTable slicers in workbook
- **Set Selection:** Filter PivotTable by slicer selection (single or multi-select)
- **Delete Slicer:** Remove PivotTable slicer

**Table Slicers:**
- **Create Table Slicer:** Add slicer for Excel Table column
- **List Table Slicers:** List all Table slicers in workbook
- **Set Table Selection:** Filter Table by slicer selection
- **Delete Table Slicer:** Remove Table slicer

**Use Cases:**
- Interactive data filtering without modifying PivotTable/Table structure
- Dashboard creation with visual filter controls
- Multi-slicer filtering for complex data analysis

---

## �🎨 Conditional Formatting (2 operations)

- **Add Rule:** Create conditional formatting rules
  - Cell value comparisons (>, <, =, etc.)
  - Expression-based formulas (custom DAX/Excel formulas)
  - Color scales, data bars, icons
- **Clear Rules:** Remove formatting from ranges

---

## 📊 Total Operations Summary

| Category | Operations |
|----------|-----------|
| File Operations | 6 |
| Power Query | 10 |
| Data Model/DAX | 18 |
| Excel Tables | 27 |
| PivotTables | 30 |
| Charts | 26 |
| Ranges | 42 |
| Worksheets | 16 |
| Connections | 9 |
| Named Ranges | 6 |
| VBA Macros | 6 |
| Slicers | 8 |
| Conditional Formatting | 2 |
| **Total** | **206** |

---

## 🚀 Key Capabilities

**Data Transformation:**
- Comprehensive Power Query M code management
- Atomic import + load workflows
- Calculated fields and members for analysis

**Data Model:**
- Full DAX measure lifecycle
- Relationship management
- Multi-table integration

**Analysis & Visualization:**
- PivotTable creation and configuration
- Chart automation
- Custom calculations

**Automation:**
- VBA macro execution and management
- Named range parameter automation
- Conditional formatting rules

**Data Loading:**
- Multiple connection type support
- OLEDB/ODBC management
- Power Query atomic workflows

---

## 🔧 Tool Selection Quick Reference

| Task | Tool |
|------|------|
| Import data | `excel_powerquery` or `excel_connection` |
| Create analysis | `excel_pivottable` (data model-based for OLAP) |
| Visualize data | `excel_chart` |
| Update parameters | `excel_namedrange` (write operation) |
| Manage formulas | `excel_range` (set-formulas) |
| Format data | `excel_range` (format-range, validate-range) |
| Script automation | `excel_vba` (run macro) |

---

## 📚 Documentation

- **[Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)** - Setup for all AI assistants
- **[MCP Server Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.McpServer/README.md)** - Tool documentation and examples
- **[CLI Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.CLI/README.md)** - Command-line reference
- **[Agent Skills](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/excel-mcp/SKILL.md)** - Cross-platform AI assistant guidance (agentskills.io)
- **[Contributing](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md)** - Development guidelines
- **[Releases](https://github.com/sbroenne/mcp-server-excel/releases)** - Latest updates and features

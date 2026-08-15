# CLI Command Reference

> Auto-generated from the built `excelcli` runtime. Use these exact command and parameter names.

### analysis

Excel what-if analysis with Goal Seek, scenarios, scenario reports, and one- or two-variable data tables. Solver is excluded because it is an optional VBA add-in that must be enabled by the user and is not part of the Excel PIA

**Actions:** `goal-seek`, `list-scenarios`, `create-scenario`, `update-scenario`, `show-scenario`, `delete-scenario`, `create-scenario-summary`, `create-data-table`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | (required) |
| `--formula-cell` | (required for: goal-seek) |
| `--goal` | (required for: goal-seek) |
| `--changing-cell` | (required for: goal-seek) |
| `--scenario-name` | (required for: create-scenario, update-scenario, show-scenario, delete-scenario) |
| `--changing-cells` | (required for: create-scenario, update-scenario) |
| `--values` | (required for: create-scenario, update-scenario) (JSON format) |
| `--comment` | Comment |
| `--locked` | Locked |
| `--hidden` | Hidden |
| `--report-type` | ReportType |
| `--result-cells` | ResultCells |
| `--table-range` | (required for: create-data-table) |
| `--row-input-cell` | RowInputCell |
| `--column-input-cell` | ColumnInputCell |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### batch

Execute multiple commands from a JSON file or stdin. Outputs NDJSON (one result per line)

| Parameter | Description |
|-----------|-------------|
| `--input` | JSON file with command array. Use '-' for stdin (NDJSON, one command per line). If omitted, reads from stdin |
| `--session` | Default session ID for all commands. Overridden by per-command sessionId. Auto-captured from session.open/create if not set |
| `--stop-on-error` | Stop execution on first error (default: continue all commands) |

### calculationmode

Control Excel recalculation (automatic vs manual). Set manual mode before bulk writes for faster performance, then recalculate once at the end

**Actions:** `get-mode`, `set-mode`, `calculate`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--mode` | Target calculation mode (required for: set-mode) |
| `--scope` | Scope: Workbook, Sheet, or Range (required for: calculate) |
| `--sheet` | Sheet name (required for Sheet/Range scope) |
| `--range` | Range address (required for Range scope) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### chart

Chart lifecycle - create, read, move, and delete embedded charts. POSITIONING (choose one): - targetRange (PREFERRED): Cell range like 'F2:K15' — positions chart within cells, no point math needed. - left/top: Manual positioning in points (72 points = 1 inch). - Neither: Auto-positions chart below all existing content (used range + other charts). COLLISION DETECTION: All create/move/fit-to-range operations automatically check for overlaps with data and other charts. Warnings are returned in the result message if collisions are detected. Always verify layout with screenshot(capture) and an explicit range that includes the chart. CHART TYPES: 70+ types available including Column, Line, Pie, Bar, Area, XY Scatter. CREATE OPTIONS: - create-from-range: Create from cell range (e.g., 'A1:D10') - create-from-table: Create from Excel Table (uses table's data range) - create-from-pivottable: Create linked PivotChart Use chartconfig for series, titles, legends, styles, placement mode

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-pivottable`, `delete`, `move`, `fit-to-range`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--chart-name` | Name of the chart (or shape name) (required for: read, delete, move, fit-to-range) |
| `--sheet` | Target worksheet name (required for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) |
| `--source-range-address` | Data range for the chart (e.g., A1:D10) (required for: create-from-range) |
| `--chart-type` | Type of chart to create (required for: create-from-range, create-from-table, create-from-pivottable) |
| `--left` | Left position in points from worksheet edge |
| `--top` | Top position in points from worksheet edge |
| `--width` | Chart width in points |
| `--height` | Chart height in points |
| `--target-range` | Cell range to position chart within (e.g., 'F2:K15'). PREFERRED over left/top. When set, left/top are ignored |
| `--table-name` | Name of the Excel Table (required for: create-from-table) |
| `--pivot-table-name` | Name of the source PivotTable (required for: create-from-pivottable) |
| `--range` | Range to fit the chart to (e.g., A1:D10) (required for: fit-to-range) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### chartconfig

Chart configuration - data source, series, type, title, axis labels, legend, and styling. SERIES MANAGEMENT: - add-series: Add data series with valuesRange (required) and optional categoryRange - remove-series: Remove series by 1-based index - set-source-range: Replace entire chart data source TITLES AND LABELS: - set-title: Set chart title (empty string hides title) - set-axis-title: Set axis labels (Category, Value, CategorySecondary, ValueSecondary) CHART STYLES: 1-48 (built-in Excel styles with different color schemes) DATA LABELS: Show values, percentages, series/category names. Positions: Center, InsideEnd, InsideBase, OutsideEnd, BestFit. TRENDLINES: Linear, Exponential, Logarithmic, Polynomial (order 2-6), Power, MovingAverage. PLACEMENT MODE: - 1: Move and size with cells - 2: Move but don't size with cells - 3: Don't move or size with cells (free floating) Use chart for lifecycle operations (create, delete, move, fit-to-range)

**Actions:** `set-source-range`, `add-series`, `remove-series`, `set-chart-type`, `set-title`, `set-axis-title`, `get-axis-number-format`, `set-axis-number-format`, `show-legend`, `set-style`, `set-placement`, `set-data-labels`, `get-axis-scale`, `set-axis-scale`, `get-gridlines`, `set-gridlines`, `set-series-format`, `set-series-chart-type`, `get-plot-options`, `set-plot-options`, `set-area-format`, `list-trendlines`, `add-trendline`, `delete-trendline`, `set-trendline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--chart-name` | Name of the chart (required) |
| `--source-range` | New data source range (e.g., Sheet1!A1:D10) (required for: set-source-range) |
| `--series-name` | Display name for the series (required for: add-series) |
| `--values-range` | Range containing series values (e.g., B2:B10) (required for: add-series) |
| `--category-range` | Optional range for category labels (e.g., A2:A10) |
| `--series-index` | 1-based index of the series to remove (required for: remove-series, set-series-format, set-series-chart-type, list-trendlines, add-trendline, delete-trendline, set-trendline) |
| `--chart-type` | New chart type to apply (required for: set-chart-type, set-series-chart-type) |
| `--title` | Title text to display (required for: set-title, set-axis-title) |
| `--axis` | Which axis to set title for (Category, Value, SeriesAxis) (required for: set-axis-title, get-axis-number-format, set-axis-number-format, get-axis-scale, set-axis-scale, set-gridlines) |
| `--number-format` | Excel number format code (e.g., "$#,##0", "0.00%") (required for: set-axis-number-format) |
| `--visible` | True to show legend, false to hide (required for: show-legend) |
| `--legend-position` | Optional position for the legend |
| `--style-id` | Excel chart style ID (1-48 for most chart types) (required for: set-style) |
| `--placement` | Placement mode: 1=MoveAndSize, 2=Move, 3=FreeFloating (required for: set-placement) |
| `--print-object` | Whether the embedded chart prints with the worksheet |
| `--locked` | Whether the embedded chart is locked when the worksheet is protected |
| `--rounded-corners` | Whether the embedded chart uses rounded corners |
| `--show-value` | Show data values on labels |
| `--show-percentage` | Show percentage values. Only meaningful for pie and doughnut chart types; setting to true on other chart types has no visual effect |
| `--show-series-name` | Show series name on labels |
| `--show-category-name` | Show category name on labels |
| `--show-bubble-size` | Show bubble size (bubble charts) |
| `--separator` | Separator string between label components |
| `--label-position` | Position of data labels relative to data points |
| `--minimum-scale` | Minimum axis value (null for auto) |
| `--maximum-scale` | Maximum axis value (null for auto) |
| `--major-unit` | Major gridline interval (null for auto) |
| `--minor-unit` | Minor gridline interval (null for auto) |
| `--show-major` | Show major gridlines (null to keep current) |
| `--show-minor` | Show minor gridlines (null to keep current) |
| `--marker-style` | Marker shape style |
| `--marker-size` | Marker size in points (2-72) |
| `--marker-background-color` | Marker fill color (#RRGGBB) |
| `--marker-foreground-color` | Marker border color (#RRGGBB) |
| `--invert-if-negative` | Invert colors for negative values |
| `--fill-color` | Series fill color as #RRGGBB |
| `--fill-transparency` | Series fill transparency from 0 (opaque) to 1 (transparent) |
| `--line-color` | Series line color as #RRGGBB |
| `--line-weight` | Series line weight in points |
| `--plot-by` | Interpret source rows or columns as data series |
| `--display-blanks-as` | How blank cells appear: gaps, zeroes, or interpolation |
| `--plot-visible-only` | True to omit hidden rows and columns |
| `--area` | Chart area or plot area (required for: set-area-format) |
| `--trendline-type` | Type of trendline (Linear, Exponential, etc.) (required for: add-trendline) |
| `--order` | Polynomial order (2-6, for Polynomial type) |
| `--period` | Moving average period (for MovingAverage type) |
| `--forward` | Periods to extend forward |
| `--backward` | Periods to extend backward |
| `--intercept` | Force trendline through specific Y-intercept |
| `--display-equation` | Display trendline equation on chart |
| `--display-r-squared` | Display R-squared value on chart |
| `--name` | Custom name for the trendline |
| `--trendline-index` | 1-based index of the trendline to delete (required for: delete-trendline, set-trendline) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### conditionalformat

Conditional formatting - visual rules based on cell values. TYPES: cellValue (requires operatorType+formula1), expression (formula only), colorScale, dataBar, iconSet, top10, aboveAverage, timePeriod, uniqueValues, blanksCondition. Both camelCase and kebab-case accepted. FORMAT: interiorColor/fontColor as #RRGGBB, fontBold/Italic, borderStyle/Color. Visual rule types use dedicated parameters on add-rule and return their type-specific configuration from list-rules / list-worksheet-rules. OPERATORS: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween. For 'between' and 'notBetween', both formula1 and formula2 are required

**Actions:** `add-rule`, `clear-rules`, `list-rules`, `list-worksheet-rules`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Sheet name (empty for active sheet) (required) |
| `--range` | Range address (A1 notation or named range) (required for: add-rule, clear-rules, list-rules) |
| `--rule-type` | Rule type: cellValue (or cell-value), expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage. Both camelCase and kebab-case accepted. (required for: add-rule) |
| `--operator-type` | Required for cellValue rules. XlFormatConditionOpe rator: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween |
| `--formula1` | Required for cellValue and expression rules. First formula/value for condition |
| `--formula2` | Required for between/notBetween cellValue rules. Second formula/value |
| `--interior-color` | Fill color (#RRGGBB or color index) |
| `--interior-pattern` | Interior pattern (1=Solid, -4142=None, 9=Gray50, etc.) |
| `--font-color` | Font color (#RRGGBB or color index) |
| `--font-bold` | Bold font |
| `--font-italic` | Italic font |
| `--border-style` | Border style: none, continuous, dash, dot, etc |
| `--border-color` | Border color (#RRGGBB or color index) |
| `--color-scale-min-type` | colorScale minimum stop type: minimum, number, percent, percentile, formula |
| `--color-scale-min-value` | colorScale minimum stop value (for number/percent/perce ntile/formula) |
| `--color-scale-min-color` | colorScale minimum stop color (#RRGGBB) |
| `--color-scale-mid-type` | colorScale midpoint stop type (supply to create a 3-color scale) |
| `--color-scale-mid-value` | colorScale midpoint stop value |
| `--color-scale-mid-color` | colorScale midpoint stop color (#RRGGBB) |
| `--color-scale-max-type` | colorScale maximum stop type: maximum, number, percent, percentile, formula |
| `--color-scale-max-value` | colorScale maximum stop value |
| `--color-scale-max-color` | colorScale maximum stop color (#RRGGBB) |
| `--data-bar-color` | dataBar fill color (#RRGGBB) |
| `--data-bar-negative-color` | dataBar negative-value bar color (#RRGGBB) |
| `--data-bar-direction` | dataBar fill direction: context, leftToRight, rightToLeft |
| `--data-bar-show-value` | dataBar show the cell value alongside the bar |
| `--data-bar-min-type` | dataBar minimum point type: automaticMinimum, minimum, number, percent, percentile, formula |
| `--data-bar-min-value` | dataBar minimum point value |
| `--data-bar-max-type` | dataBar maximum point type: automaticMaximum, maximum, number, percent, percentile, formula |
| `--data-bar-max-value` | dataBar maximum point value |
| `--icon-set-id` | iconSet id: 3Arrows, 3TrafficLights1, 4Ratings, 5Quarters, etc |
| `--icon-set-reverse` | iconSet reverse icon order |
| `--icon-set-show-icon-only` | iconSet show only the icon (hide the value) |
| `--icon-threshold1-type` | iconSet threshold 1 type: percent, number, percentile, formula |
| `--icon-threshold1-value` | iconSet threshold 1 value |
| `--icon-threshold2-type` | iconSet threshold 2 type |
| `--icon-threshold2-value` | iconSet threshold 2 value |
| `--icon-threshold3-type` | iconSet threshold 3 type |
| `--icon-threshold3-value` | iconSet threshold 3 value |
| `--icon-threshold4-type` | iconSet threshold 4 type |
| `--icon-threshold4-value` | iconSet threshold 4 value |
| `--rank` | top10 rank (number of values, or percent when top10Percent is true) |
| `--top10-percent` | top10 treat rank as a percentage |
| `--top-bottom` | top10 direction: top or bottom |
| `--above-below` | aboveAverage selector: aboveAverage, belowAverage, aboveStdDev, belowStdDev, equalAboveAverage, equalBelowAverage |
| `--date-period` | timePeriod period: today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### connection

Data connections (OLEDB, ODBC, ODC import). TEXT/WEB/CSV: Use querytable for direct local imports or powerquery for transformations. Power Query connections auto-redirect to powerquery. TIMEOUT: 30 min auto-timeout for refresh/load-to

**Actions:** `list`, `view`, `create`, `refresh`, `get-refresh-status`, `cancel-refresh`, `delete`, `load-to`, `get-properties`, `set-properties`, `test`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--connection-name` | Name of the connection to view (required for: view, create, refresh, get-refresh-status, cancel-refresh, delete, load-to, get-properties, set-properties, test) |
| `--connection-string` | OLEDB or ODBC connection string (required for: create) |
| `--command-text` | SQL query or table name |
| `--description` | Optional description for the connection |
| `--timeout` | Optional timeout for the refresh operation Accepts seconds (e.g. 600) or TimeSpan format (e.g. 00:10:00) |
| `--sheet` | Target worksheet name (required for: load-to) |
| `--background-query` | Run query in background (null to keep current) |
| `--refresh-on-file-open` | Refresh when file opens (null to keep current) |
| `--save-password` | Save password in connection (null to keep current) |
| `--refresh-period` | Auto-refresh interval in minutes (null to keep current) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### datamodel

Data Model (Power Pivot) - DAX measures and table management. CRITICAL: WORKSHEET TABLES AND DATA MODEL ARE SEPARATE! - After table append changes, Data Model still has OLD data - MUST call refresh to sync changes - Power Query refresh auto-syncs (no manual refresh needed) PREREQUISITE: Tables must be added to the Data Model first. Use table add-to-datamodel for worksheet tables, or powerquery to import and load data directly to the Data Model. DAX MEASURES: - Create with DAX formulas like 'SUM(Sales[Amount])' - DAX formulas are preserved exactly by default - Set formatDax=true only with user consent; it sends formulas to daxformatter.com - Read operations return raw DAX as stored DAX EVALUATE QUERIES: - Use evaluate to execute DAX EVALUATE queries against the Data Model - Returns tabular results from queries like 'EVALUATE TableName' - Supports complex DAX: SUMMARIZE, FILTER, CALCULATETABLE, TOPN, etc. DMV (DYNAMIC MANAGEMENT VIEW) QUERIES: - Use execute-dmv to query Data Model metadata via SQL-like syntax - Syntax: SELECT * FROM $SYSTEM.SchemaRowset (ONLY SELECT * supported) - Use DISCOVER_SCHEMA_ROWSETS to list all available DMVs Use datamodelrel for relationships between tables

**Actions:** `list-tables`, `list-columns`, `read-table`, `read-info`, `read-connection`, `list-measures`, `read`, `delete-measure`, `delete-table`, `rename-table`, `refresh`, `create-measure`, `update-measure`, `evaluate`, `execute-dmv`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--table-name` | Name of the table to list columns from (required for: list-columns, read-table, delete-table, create-measure) |
| `--measure-name` | Name of the measure to get (required for: read, delete-measure, create-measure, update-measure) |
| `--old-name` | Current name of the table (required for: rename-table) |
| `--new-name` | New name for the table (required for: rename-table) |
| `--timeout` | Optional: Timeout for the refresh operation Accepts seconds (e.g. 600) or TimeSpan format (e.g. 00:10:00) |
| `--dax-formula` | DAX formula for the measure (required for: create-measure) |
| `--dax-formula-file` | Path to file containing daxFormula |
| `--format-type` | Optional: Format type (Currency, Decimal, Percentage, General) |
| `--description` | Optional: Description of the measure |
| `--format-dax` | Whether to send the DAX formula to the remote daxformatter.com service before saving. Defaults to false to preserve privacy |
| `--dax-query` | DAX EVALUATE query (e.g., "EVALUATE 'TableName'" or "EVALUATE SUMMARIZE(...)") (required for: evaluate) |
| `--dax-query-file` | Path to file containing daxQuery |
| `--dmv-query` | DMV query in SQL-like syntax (e.g., "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES") (required for: execute-dmv) |
| `--dmv-query-file` | Path to file containing dmvQuery |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### datamodelrelationship

Data Model relationships - link tables for cross-table DAX calculations. CRITICAL: Deleting or recreating tables removes ALL their relationships. Use list-relationships before table operations to backup, then recreate relationships after schema changes. RELATIONSHIP REQUIREMENTS: - Both tables must exist in the Data Model first - Columns must have compatible data types - fromTable/fromColumn = many-side (detail table, foreign key) - toTable/toColumn = one-side (lookup table, primary key) ACTIVE VS INACTIVE: - Only ONE active relationship can exist between two tables - Use active=false when creating alternative paths - DAX USERELATIONSHIP() activates inactive relationships

**Actions:** `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--from-table` | Source table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--from-column` | Source column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-table` | Target table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-column` | Target column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--active` | Whether the relationship should be active (default: true) (required for: update-relationship) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### diag

Diagnostic commands for testing CLI/MCP infrastructure without Excel. These commands validate parameter parsing, routing, JSON serialization, and error handling — no Excel COM session needed

**Actions:** `ping`, `echo`, `validate-params`

| Parameter | Description |
|-----------|-------------|
| `--message` | The message to echo back (required) (required for: echo) |
| `--tag` | Optional tag to include in the response |
| `--name` | Required name parameter (required for: validate-params) |
| `--count` | Required integer parameter (required for: validate-params) |
| `--label` | Optional label parameter |
| `--verbose` | Optional boolean flag (default: false) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### drawing

Worksheet drawing objects and sparklines. OBJECTS: list/read/update/delete images, AutoShapes, text boxes, connectors, and worksheet Forms controls. SHAPE TYPES: common geometric, arrow, and flowchart AutoShapes. FORMATTING: geometry, text, fill/line/font colors, rotation, visibility, locking, placement, and alternative text. FORMS CONTROLS: safe worksheet Forms controls only. linkedCell applies to CheckBox, DropDown, ListBox, OptionButton, ScrollBar, and Spinner; inputRange applies only to DropDown and ListBox. ActiveX/OLE controls and macro assignment are intentionally excluded. SPARKLINES: list/read/create/update/delete line, column, and win/loss sparklines. COLORS: use #RRGGBB hexadecimal values

**Actions:** `list-objects`, `get-object`, `add-image`, `add-shape`, `add-text-box`, `add-connector`, `add-form-control`, `update-object`, `delete-object`, `list-sparklines`, `get-sparkline`, `add-sparkline`, `update-sparkline`, `delete-sparkline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | (required) |
| `--object-name` | (required for: get-object, update-object, delete-object) |
| `--image-path` | (required for: add-image) |
| `--name` | Name |
| `--left` | Left |
| `--top` | Top |
| `--width` | Width |
| `--height` | Height |
| `--lock-aspect-ratio` | LockAspectRatio |
| `--shape-type` | ShapeType |
| `--text` | (required for: add-text-box) |
| `--fill-color` | FillColor |
| `--line-color` | LineColor |
| `--line-weight` | LineWeight |
| `--font-size` | FontSize |
| `--font-color` | FontColor |
| `--connector-type` | ConnectorType |
| `--begin-x` | BeginX |
| `--begin-y` | BeginY |
| `--end-x` | EndX |
| `--end-y` | EndY |
| `--control-type` | ControlType |
| `--linked-cell` | LinkedCell |
| `--input-range` | InputRange |
| `--new-name` | New name (for rename) |
| `--rotation` | Rotation |
| `--visible` | Visible |
| `--locked` | Locked |
| `--placement` | Placement |
| `--alternative-text` | AlternativeText |
| `--location-range` | (required for: get-sparkline, add-sparkline, update-sparkline, delete-sparkline) |
| `--source-range` | (required for: add-sparkline) |
| `--sparkline-type` | SparklineType |
| `--show-markers` | ShowMarkers |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### namedrange

Named ranges for formulas/parameters. LIST: returns visible user-defined names; hidden/internal Excel names are omitted before value inspection, and large ranges return metadata without materializing values. CREATE/UPDATE: value is cell reference (e.g., 'Sheet1!$A$1'). WRITE: value is data to store. TIP: use range get-values/set-values with the named range as the range address for bulk data read/write

**Actions:** `list`, `write`, `read`, `update`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--name` | Name of the named range (required for: write, read, update, create, delete) |
| `--value` | Value to set (required for: write) |
| `--reference` | New cell reference (e.g., Sheet1!$A$1:$B$10) (required for: update, create) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pivottable

PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete. Use pivottablefield for field operations, pivottablecalc for calculated fields and layout. BEST PRACTICE: Use 'list' before creating. Prefer 'refresh' or field modifications over delete+recreate. Delete+recreate loses field configurations, filters, sorting, and custom layouts. REFRESH: Call 'refresh' after configuring fields with pivottablefield to update the visual display. This is especially important for OLAP/Data Model PivotTables where field operations are structural only and don't automatically trigger a visual refresh. CREATE OPTIONS: - 'create-from-range': Use source sheet and range address for data range - 'create-from-table': Use an Excel Table (ListObject) as source - 'create-from-datamodel': Use a Power Pivot Data Model table as source

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-datamodel`, `delete`, `refresh`, `get-cache-options`, `set-cache-options`, `drill-through`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh, get-cache-options, set-cache-options, drill-through) |
| `--source-sheet` | Source worksheet name (required for: create-from-range) |
| `--source-range` | Source range address (e.g., "A1:F100") (required for: create-from-range) |
| `--destination-sheet` | Destination worksheet name (required for: create-from-range, create-from-table, create-from-datamodel) |
| `--destination-cell` | Destination cell address (e.g., "A1") (required for: create-from-range, create-from-table, create-from-datamodel) |
| `--table-name` | Name of the Excel Table (required for: create-from-table, create-from-datamodel) |
| `--timeout` | Optional timeout for the refresh operation Accepts seconds (e.g. 600) or TimeSpan format (e.g. 00:10:00) |
| `--enable-refresh` | Allow the PivotCache to refresh |
| `--refresh-on-file-open` | Refresh the PivotCache when the workbook opens |
| `--missing-items-limit` | How many deleted source items Excel retains in the cache |
| `--optimize-cache` | Optimize the cache when it is constructed |
| `--save-source-data` | Save source records with the PivotTable |
| `--cell-address` | Value-cell address on the PivotTable worksheet (for example, G4) (required for: drill-through) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pivottablecalc

PivotTable calculated fields/members, layout configuration, and data extraction. Use pivottable for lifecycle, pivottablefield for field placement. CALCULATED FIELDS (for regular PivotTables): - Create custom fields using formulas like '=Revenue-Cost' or '=Quantity*UnitPrice' - Can reference existing fields by name - After creating, use pivottablefield add-value-field to add to Values area - For complex multi-table calculations, prefer DAX measures with datamodel CALCULATED MEMBERS (for OLAP/Data Model PivotTables only): - Create using MDX expressions - Member types: Member, Set, Measure LAYOUT OPTIONS: - 0 = Compact (default, fields in single column) - 1 = Tabular (each field in separate column - best for export/analysis) - 2 = Outline (hierarchical with expand/collapse)

**Actions:** `get-data`, `create-calculated-field`, `list-calculated-fields`, `delete-calculated-field`, `list-calculated-members`, `create-calculated-member`, `delete-calculated-member`, `set-layout`, `set-subtotals`, `set-grand-totals`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required) |
| `--field-name` | Name for the calculated field (required for: create-calculated-field, delete-calculated-field, set-subtotals) |
| `--formula` | Formula using field references (e.g., "=Revenue-Cost") (required for: create-calculated-field, create-calculated-member) |
| `--member-name` | Name for the calculated member (MDX naming format) (required for: create-calculated-member, delete-calculated-member) |
| `--type` | Type of calculated member (Member, Set, or Measure) |
| `--solve-order` | Solve order for calculation precedence (default: 0) |
| `--display-folder` | Display folder path for organizing measures (optional) |
| `--number-format` | Number format code for the calculated member (optional) |
| `--row-layout` | Layout form: 0=Compact, 1=Tabular, 2=Outline (required for: set-layout) |
| `--show-subtotals` | True to show automatic subtotals, false to hide (required for: set-subtotals) |
| `--show-row-grand-totals` | Show row grand totals (bottom summary row) (required for: set-grand-totals) |
| `--show-column-grand-totals` | Show column grand totals (right summary column) (required for: set-grand-totals) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pivottablefield

PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping. Use pivottable for lifecycle, pivottablecalc for calculated fields and layout. IMPORTANT: Field operations modify structure only. Call pivottable refresh after configuring fields to update the visual display, especially for OLAP/Data Model PivotTables. FIELD AREAS: - Row fields: Group data by categories (add-row-field) - Column fields: Create column headers (add-column-field) - Value fields: Aggregate numeric data with Sum, Count, Average, etc. (add-value-field) - Filter fields: Add report-level filters (add-filter-field) AGGREGATION FUNCTIONS: Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP GROUPING: - Date fields: Group by Days, Months, Quarters, Years (group-by-date) - Numeric fields: Group by ranges with start/end/interval (group-by-numeric) NUMBER FORMAT: Use US format codes like '#,##0.00' for currency or '0.00%' for percentages

**Actions:** `list-fields`, `add-row-field`, `add-column-field`, `add-value-field`, `add-filter-field`, `remove-field`, `set-field-function`, `set-field-name`, `set-field-format`, `set-field-filter`, `sort-field`, `group-by-date`, `group-by-numeric`, `group-items`, `ungroup-field`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required) |
| `--field-name` | Name of the field to add (required for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, group-by-date, group-by-numeric, group-items) |
| `--position` | Optional position in row area (1-based) |
| `--aggregation-function` | Aggregation function (for Regular and OLAP auto-create mode) (required for: set-field-function) |
| `--custom-name` | Optional custom name for the field/measure (required for: set-field-name) |
| `--number-format` | Number format string (required for: set-field-format) |
| `--selected-values` | Values to show (others will be hidden) (required for: set-field-filter) (JSON format) |
| `--direction` | Sort direction |
| `--interval` | Grouping interval (Months, Quarters, Years) (required for: group-by-date) |
| `--start` | Starting value (null = use field minimum) |
| `--end-value` | Ending value (null = use field maximum) |
| `--interval-size` | Size of each group (e.g., 100 for groups of 100) (required for: group-by-numeric) |
| `--item-names` | Exact item captions to group (required for: group-items) (JSON format) |
| `--group-name` | Caption for the new group (required for: group-items) |
| `--grouped-field-name` | Generated grouped field name (required for: ungroup-field) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### powerquery

Power Query M code and data loading. TEST-FIRST DEVELOPMENT WORKFLOW (BEST PRACTICE): 1. evaluate - Test M code WITHOUT persisting (catches syntax errors, validates sources, shows data preview) 2. create/update - Store VALIDATED query in workbook 3. refresh/load-to - Load data to destination Skip evaluate only for trivial literal tables. IF CREATE/UPDATE FAILS: Use evaluate to get the actual M engine error message, fix code, retry. DATETIME COLUMNS: Always include Table.TransformColumnTypes() in M code to set column types explicitly. Without explicit types, dates may be stored as numbers and Data Model relationships may fail. DESTINATIONS: 'worksheet' (default), 'data-model' (for DAX), 'both', 'connection-only'. Use 'data-model' to load to Power Pivot, then use datamodel to create DAX measures. M-CODE: Preserved exactly by default. Set formatMCode=true only with explicit user consent; it sends M code to powerqueryformatter.com. TARGET CELL: targetCellAddress places tables without clearing sheet. TIMEOUT: 30 min auto-timeout for refresh and load-to. For quick queries, use timeout=60 or similar. timeout=0 or omitted uses the 30 min default

**Actions:** `list`, `view`, `refresh`, `get-load-config`, `delete`, `create`, `update`, `load-to`, `refresh-all`, `rename`, `unload`, `evaluate`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--query-name` | Name of the query to view (required for: view, refresh, get-load-config, delete, create, update, load-to, unload) |
| `--timeout` | Maximum time to wait for refresh (required for: refresh) Accepts seconds (e.g. 600) or TimeSpan format (e.g. 00:10:00) |
| `--m-code` | Raw M code (inline string) (required for: create, update, evaluate) |
| `--m-code-file` | Path to file containing mCode |
| `--load-destination` | Load destination mode (required for: load-to) |
| `--target-sheet` | Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted) |
| `--target-cell-address` | Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data |
| `--format-m-code` | Whether to send M code to the remote powerqueryformatter.com service before saving. Defaults to false to preserve privacy |
| `--refresh` | Whether to refresh data after update (default: true) |
| `--old-name` | Current name of the query (required for: rename) |
| `--new-name` | New name for the query (required for: rename) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pythoninexcel

Microsoft 365 "Python in Excel" (=PY()) formulas. REQUIRES: a real Excel session signed into a licensed Microsoft 365 account with Python in Excel enabled, plus internet access — the Python code executes in a Microsoft-hosted cloud sandbox, not locally. Not available offline or with perpetual-license Excel. SET-FORMULA writes '=PY("<code>", returnType)' via Range.Formula2. returnType 0 = "Excel Value" (a plain value/array), 1 = "Python Object" (a rich data type card, e.g. a DataFrame). The returnType argument must always be passed explicitly — omitting it causes a #NAME? error. GET-RESULT reads the computed value back, polling until the cloud round-trip finishes (a fresh formula reads as #BUSY! while the cloud Python sandbox is still computing). Completion is detected deterministically from Excel's calculation state plus a per-cell #BUSY! guard, so a converged result is not confused with the in-flight #BUSY! placeholder. If the cloud backend is still busy at the deadline (e.g. a cold start), GET-RESULT reports that and asks the caller to retry or raise maxWaitSeconds rather than returning a placeholder. If Excel evaluates PY() as #NAME?, both actions report that Python in Excel is unavailable in the current session instead of returning the raw worksheet error. DATA BINDING: reference live worksheet data inside the Python code with xl("A1:A6"), xl("Sheet1!A1:A6"), or a named range xl("MyRange") — all work when the formula is authored via this tool, the same as if typed interactively. TIP: xl() returns a pandas DataFrame/Series (not a plain list) unless you pass headers explicitly; prefer pandas methods (.sum()/.mean()/.max()) over Python builtins (sum()/len()) to avoid getting a Series back instead of a scalar total

**Actions:** `set-formula`, `get-result`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet name (required) |
| `--range` | Target cell address (e.g. "D1") (required) |
| `--code` | Python source code (e.g. "xl('A1:A6').sum()") (required for: set-formula) |
| `--return-type` | 0 = Excel Value (default), 1 = Python Object |
| `--max-wait-seconds` | Maximum seconds to poll for the cloud result before giving up (default: 30). Must be shorter than the session operation timeout. Returns as soon as the result is ready |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### querytable

Worksheet QueryTable lifecycle and configuration for local COM text, CSV, and legacy web imports. Use powerquery for modern connectors and transformations

**Actions:** `list`, `view`, `create-text`, `create-web`, `set-properties`, `refresh`, `get-refresh-status`, `cancel-refresh`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | (required for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) |
| `--query-table-name` | (required for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) |
| `--source-path` | (required for: create-text) |
| `--destination-address` | (required for: create-text, create-web) |
| `--delimiter` | Delimiter |
| `--text-qualifier` | TextQualifier |
| `--encoding` | Encoding |
| `--has-headers` | HasHeaders |
| `--url` | (required for: create-web) |
| `--selection-type` | SelectionType |
| `--web-tables` | WebTables |
| `--formatting` | Formatting |
| `--background-query` | BackgroundQuery |
| `--refresh-on-file-open` | RefreshOnFileOpen |
| `--refresh-period` | RefreshPeriod |
| `--adjust-column-width` | AdjustColumnWidth |
| `--preserve-formatting` | PreserveFormatting |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### range

Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions. Use rangeedit for insert/delete/find/sort. Use rangeformat for styling/validation. Use rangelink for hyperlinks and cell protection. Calculation mode and explicit recalculation are handled by calculationmode. BEST PRACTICE: Use 'get-values' to check existing data before overwriting. Use 'clear-contents' (not 'clear-all') to preserve cell formatting when clearing data. set-values preserves existing formatting; use set-number-format after if format change needed. DATA FORMAT: values and formulas are 2D JSON arrays representing rows and columns. Example: [[row1col1, row1col2], [row2col1, row2col2]] Single cell returns [[value]] (always 2D). REQUIRED PARAMETERS: - sheetName + rangeAddress for cell operations (e.g., sheetName='Sheet1', rangeAddress='A1:D10') - For named ranges, use sheetName='' (empty string) and rangeAddress='MyNamedRange' COPY OPERATIONS: Specify source and target sheet/range for copy operations. NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', 'mm/dd/yyyy', '0.00%')

**Actions:** `get-values`, `set-values`, `get-formulas`, `set-formulas`, `validate-formulas`, `clear-all`, `clear-contents`, `clear-formats`, `copy`, `copy-values`, `copy-formulas`, `get-number-formats`, `set-number-format`, `set-number-formats`, `get-used-range`, `get-current-region`, `get-info`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only (required for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info) |
| `--range` | Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name (e.g., 'SalesData') (required for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info) |
| `--values` | 2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Optional if valuesFile is provided. (JSON format) |
| `--values-file` | Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter |
| `--formulas` | 2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Optional if formulasFile is provided. (JSON format) |
| `--formulas-file` | Path to a JSON file containing the formulas as a 2D array. Alternative to inline formulas parameter |
| `--source-sheet` | Source worksheet name for copy operations (required for: copy, copy-values, copy-formulas) |
| `--source-range` | Source range address for copy operations (e.g., 'A1:D10') (required for: copy, copy-values, copy-formulas) |
| `--target-sheet` | Target worksheet name for copy operations (required for: copy, copy-values, copy-formulas) |
| `--target-range` | Target range address - can be single cell for paste destination (e.g., 'A1') (required for: copy, copy-values, copy-formulas) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00' for numbers, 'mm/dd/yyyy' for dates, '0.00%' for percentages, 'General' for default, '@' for text) (required for: set-number-format) |
| `--formats` | 2D array of format codes - same dimensions as target range (e.g., [['#,##0.00', '0.00%'], ['mm/dd/yyyy', 'General']]). Optional if formatsFile is provided. (JSON format) |
| `--formats-file` | Path to a JSON file containing 2D array of format codes. Alternative to inline formats parameter |
| `--cell-address` | Single cell address (e.g., 'B5') - expands to contiguous data region around this cell (required for: get-current-region) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### rangeedit

Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data. Use range for values/formulas/copy/clear operations. INSERT/DELETE CELLS: Specify shift direction to control how surrounding cells move. - Insert: 'Down' or 'Right' - Delete: 'Up' or 'Left' INSERT/DELETE ROWS: Use row range like '5:10' to insert/delete rows 5-10. INSERT/DELETE COLUMNS: Use column range like 'B:D' to insert/delete columns B-D. FIND/REPLACE: Search within the specified range with optional case/cell matching. - Find returns up to 10 matching cell addresses with total count. - Replace modifies all matches by default. SORT: Specify sortColumns as array of {columnIndex: 1, ascending: true} objects. Column indices are 1-based relative to the range

**Actions:** `insert-cells`, `delete-cells`, `insert-rows`, `delete-rows`, `insert-columns`, `delete-columns`, `find`, `replace`, `sort`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range (required) |
| `--range` | Cell range address where cells will be inserted (e.g., 'A1:D10') (required) |
| `--insert-shift` | Direction to shift existing cells: 'Down' or 'Right' (required for: insert-cells) |
| `--delete-shift` | Direction to shift remaining cells: 'Up' or 'Left' (required for: delete-cells) |
| `--search-value` | Text or value to search for (required for: find) |
| `--find-options` | Search options: matchCase (default: false), matchEntireCell (default: false), searchFormulas (default: true) (required for: find) |
| `--find-value` | Text or value to search for (required for: replace) |
| `--replace-value` | Text or value to replace matches with (required for: replace) |
| `--replace-options` | Replace options: matchCase (default: false), matchEntireCell (default: false), replaceAll (default: true) (required for: replace) |
| `--sort-columns` | Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range (required for: sort) (JSON format) |
| `--has-headers` | Whether the range has a header row to exclude from sorting (default: true) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### rangeformat

Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions. Use range tool for values/formulas/copy/clear operations. set-style: Apply a named Excel style (Heading 1, Good, Bad, Neutral, Normal). Best for semantic status labels (Good/Bad/Neutral have fill colours and are theme-aware) and document hierarchy (Heading 1/2/3). NOTE: Heading styles do NOT apply a fill colour — use format-range when you need a coloured header row. format-range: Apply any combination of bold, fillColor, fontColor, alignment, borders. Required whenever you need a fill colour or custom branding. Pass ALL desired properties in a SINGLE call — do not call format-range multiple times for the same range. format-ranges: Apply one shared formatting payload to multiple ranges on the same worksheet. Prefer this over repeated format-range calls when the same styling applies to multiple non-contiguous targets. All target ranges are validated before formatting begins. If any target range is invalid, nothing is formatted. COLORS: Hex '#RRGGBB' (e.g., '#FF0000' for red, '#00FF00' for green) FONT: size in points (e.g., 12, 14, 16), alignment: 'left', 'center', 'right' / 'top', 'middle', 'bottom' DATA VALIDATION: Restrict cell input with validation rules: - Types: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom' - For list validation, formula1 is the list source (e.g., '=$A$1:$A$10' or '"Option1,Option2,Option3"') - Operators: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual' MERGE: Combines cells into one. Only top-left cell value is preserved

**Actions:** `set-style`, `get-style`, `format-range`, `format-ranges`, `validate-range`, `get-validation`, `remove-validation`, `auto-fit-columns`, `auto-fit-rows`, `merge-cells`, `unmerge-cells`, `get-merge-info`, `set-column-width`, `set-row-height`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range (required) |
| `--range` | Cell range address (e.g., 'A1:D10') (required for: set-style, get-style, format-range, validate-range, get-validation, remove-validation, auto-fit-columns, auto-fit-rows, merge-cells, unmerge-cells, get-merge-info, set-column-width, set-row-height) |
| `--style-name` | Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent'). Use 'Normal' to reset. (required for: set-style) |
| `--font-name` | Font family name (e.g., 'Arial', 'Calibri', 'Times New Roman') |
| `--font-size` | Font size in points (e.g., 10, 11, 12, 14, 16) |
| `--bold` | Whether to apply bold formatting |
| `--italic` | Whether to apply italic formatting |
| `--underline` | Whether to apply underline formatting |
| `--font-color` | Font (foreground) color as hex '#RRGGBB' (e.g., '#FF0000' for red) |
| `--fill-color` | Cell fill (background) color as hex '#RRGGBB' (e.g., '#FFFF00' for yellow) |
| `--border-style` | Border line style: 'continuous', 'dash', 'dot', 'dashdot', 'dashdotdot', 'double', 'slantdashdot', 'none' |
| `--border-color` | Border color as hex '#RRGGBB' |
| `--border-weight` | Border weight: 'hairline', 'thin', 'medium', 'thick' |
| `--horizontal-alignment` | Horizontal text alignment: 'left', 'center', 'right', 'justify', 'fill' |
| `--vertical-alignment` | Vertical text alignment: 'top', 'center' (or 'middle'), 'bottom', 'justify' |
| `--wrap-text` | Whether to wrap text within cells |
| `--orientation` | Text rotation in degrees (-90 to 90, or 255 for vertical) |
| `--range-addresses` | Cell range addresses to format (e.g., 'A1:D1', 'A3:D3') (required for: format-ranges) |
| `--number-format` | Excel number format code applied to all target ranges (e.g., '0.00%' for percentage, '$#,##0.00' for currency, 'm/d/yyyy' for date). LLMs know Excel format codes natively |
| `--validation-type` | Data validation type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom' (required for: validate-range) |
| `--validation-operator` | Validation comparison operator: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual' |
| `--formula1` | First validation formula/value - for list validation use range '=$A$1:$A$10' or inline '"A,B,C"' |
| `--formula2` | Second validation formula/value - required only for 'between' and 'notBetween' operators |
| `--show-input-message` | Whether to show input message when cell is selected (default: false) |
| `--input-title` | Title for the input message popup |
| `--input-message` | Text for the input message popup |
| `--show-error-alert` | Whether to show error alert on invalid input (default: true) |
| `--error-style` | Error alert style: 'stop' (prevents entry), 'warning' (allows override), 'information' (allows entry) |
| `--error-title` | Title for the error alert popup |
| `--error-message` | Text for the error alert popup |
| `--ignore-blank` | Whether to allow blank cells in validation (default: true) |
| `--show-dropdown` | Whether to show dropdown arrow for list validation (default: true) |
| `--column-width` | Width in points (1 point = 1/72 inch, approx 0.35mm). Standard width ~8.43 points. Range: 0.25-409 points. (required for: set-column-width) |
| `--row-height` | Height in points (1 point = 1/72 inch, approx 0.35mm). Default row height ~15 points. Range: 0-409 points. (required for: set-row-height) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### rangelink

Hyperlink, threaded comment, and cell protection operations for Excel ranges. Use range for values/formulas, rangeformat for styling. HYPERLINKS: - 'add-hyperlink': Add an external or internal workbook hyperlink - 'update-hyperlink': Update the target or display metadata of an existing hyperlink - 'remove-hyperlink': Remove hyperlink(s) from cells while keeping the cell content - 'list-hyperlinks': Get all hyperlinks on a worksheet - 'get-hyperlink': Get hyperlink details for a specific cell CELL PROTECTION: - 'set-cell-lock': Lock or unlock cells (only effective when sheet protection is enabled) - 'get-cell-lock': Check if cells are locked Note: Cell locking only takes effect when the worksheet is protected

**Actions:** `add-hyperlink`, `update-hyperlink`, `remove-hyperlink`, `list-hyperlinks`, `get-hyperlink`, `add-threaded-comment`, `list-threaded-comments`, `add-threaded-comment-reply`, `delete-threaded-comment`, `set-cell-lock`, `get-cell-lock`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet (required) |
| `--cell-address` | Single cell address (e.g., 'A1') (required for: add-hyperlink, update-hyperlink, get-hyperlink, add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment) |
| `--url` | Optional external URL or file path. Omit for an internal workbook link |
| `--display-text` | Text to display in the cell (optional, defaults to URL) |
| `--tooltip` | Tooltip text shown on hover (optional) |
| `--sub-address` | Optional internal workbook target such as "'Sheet2'!A1" |
| `--range` | Cell range address to remove hyperlinks from (e.g., 'A1:D10') (required for: remove-hyperlink, set-cell-lock, get-cell-lock) |
| `--text` | (required for: add-threaded-comment, add-threaded-comment-reply) |
| `--locked` | Lock status: true = locked (protected when sheet protection enabled), false = unlocked (editable) (required for: set-cell-lock) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### screenshot

Capture Excel worksheet content as images for visual verification. Uses Excel's built-in rendering (CopyPicture) to capture ranges as PNG images. Captures formatting, conditional formatting, charts, and all visual elements. ACTIONS: - capture: Capture a specific range as an image - capture-sheet: Capture the worksheet's used cell range. For chart-only sheets or charts beyond used cells, use capture with an explicit range. RETURNS: Base64-encoded image data with dimensions metadata. For MCP: returned as native ImageContent (no file handling needed). For CLI: use --output <path> to save the image directly to a PNG/JPEG file instead of returning base64 inline. Quality defaults to Medium (JPEG 75% scale) which is 4-8x smaller than High (PNG). Use High only when fine detail inspection is needed

**Actions:** `capture`, `capture-sheet`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet name (null for active sheet) |
| `--range` | Range to capture (e.g., "A1:F20") |
| `--quality` | Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### service

Service lifecycle management: start, stop, status

#### service start

Start the ExcelMCP Service if not already running

#### service stop

Gracefully stop the ExcelMCP Service

#### service status

Show service status (running, PID, sessions, uptime)

### session

Session management. WORKFLOW: open -> use sessionId -> close (--save to persist). Use --show for IRM/auth prompts

#### session create

Create a new Excel file, open it, and create a session. Add --show for visible Excel

| Parameter | Description |
|-----------|-------------|
| `<file>` | Path to the new Excel file to create |
| `--timeout` | Session open/create and operation timeout in seconds (default: 120) |
| `--show` | Show the Excel window for IRM/auth prompts instead of running hidden |

#### session open

Open an Excel file and create a session. Add --show for visible Excel

| Parameter | Description |
|-----------|-------------|
| `<file>` | Path to the Excel file to open |
| `--timeout` | Session open and operation timeout in seconds (default: 120) |
| `--show` | Show the Excel window for IRM/auth prompts instead of running hidden |

#### session close

Close a session. Use --save to persist changes

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID to close |
| `--save` | Save changes before closing |

#### session list

List active sessions

#### session save

Save a session without closing it

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID to save |

### sheet

Worksheet lifecycle management: create, rename, copy, delete, move, list sheets. Use range for data operations. Use sheetstyle for tab colors and visibility. ATOMIC OPERATIONS: 'copy-to-file' and 'move-to-file' don't require a session - they open/close files automatically. POSITIONING: For 'move', 'copy-to-file', 'move-to-file' - use 'before' OR 'after' (not both) to position the sheet relative to another. If neither specified, moves to end. NOTE: MCP tool is manually implemented in ExcelWorksheetTool.cs to properly handle mixed session requirements (copy-to-file and move-to-file are atomic and don't need sessions)

**Actions:** `list`, `create`, `rename`, `copy`, `delete`, `move`, `copy-to-file`, `move-to-file`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--file-path` | Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook |
| `--sheet` | Name for the new worksheet (required for: create, delete, move) |
| `--old-name` | Current name of the worksheet (required for: rename) |
| `--new-name` | New name for the worksheet (required for: rename) |
| `--source-name` | Name of the source worksheet (required for: copy) |
| `--target-name` | Name for the copied worksheet (required for: copy) |
| `--before-sheet` | Optional: Name of sheet to position before |
| `--after-sheet` | Optional: Name of sheet to position after |
| `--source-file` | Full path to the source workbook (required for: copy-to-file, move-to-file) |
| `--source-sheet` | Name of the sheet to copy (required for: copy-to-file, move-to-file) |
| `--target-file` | Full path to the target workbook (required for: copy-to-file, move-to-file) |
| `--target-sheet-name` | Optional: New name for the copied sheet (default: keeps original name) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### slicer

Slicer visual filters for PivotTables and Excel Tables. PIVOTTABLE SLICERS: create-slicer, list-slicers, set-slicer-selection, delete-slicer. TABLE SLICERS: create-table-slicer, list-table-slicers, set-table-slicer-selection, delete-table-slicer. NAMING: Auto-generate descriptive names like {FieldName}Slicer (e.g., RegionSlicer). SELECTION: selectedItems as list of strings. Empty list clears filter (shows all items). Set clearFirst=false to add to existing selection

**Actions:** `create-slicer`, `list-slicers`, `set-slicer-selection`, `delete-slicer`, `create-table-slicer`, `list-table-slicers`, `set-table-slicer-selection`, `delete-table-slicer`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable to create slicer for (required for: create-slicer) |
| `--field-name` | Name of the field to use for the slicer (required for: create-slicer) |
| `--slicer-name` | Name for the new slicer (required for: create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) |
| `--destination-sheet` | Worksheet where slicer will be placed (required for: create-slicer, create-table-slicer) |
| `--position` | Top-left cell position for the slicer (e.g., "H2") (required for: create-slicer, create-table-slicer) |
| `--selected-items` | Items to select (show in PivotTable) (required for: set-slicer-selection, set-table-slicer-selection) (JSON format) |
| `--clear-first` | If true, clears existing selection before setting new items (default: true) |
| `--table-name` | Name of the Excel Table (required for: create-table-slicer) |
| `--column-name` | Name of the column to use for the slicer (required for: create-table-slicer) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### table

Excel Tables (ListObjects) - lifecycle and data operations. Tables provide structured references, automatic formatting, and Data Model integration. BEST PRACTICE: Use 'list' to check existing tables before creating. Prefer 'append'/'resize'/'rename' over delete+recreate to preserve references. WARNING: Deleting tables used as PivotTable sources or in Data Model relationships will break those objects. DATA MODEL WORKFLOW: To analyze worksheet data with DAX/Power Pivot: 1. Create or identify an Excel Table on a worksheet 2. Use 'add-to-datamodel' to add the table to Power Pivot 3. Then use datamodel to create DAX measures on it DAX-BACKED TABLES: Create tables populated by DAX EVALUATE queries: - 'create-from-dax': Create a new table backed by a DAX query (e.g., SUMMARIZE, FILTER) - 'update-dax': Update the DAX query for an existing DAX-backed table - 'get-dax': Get the DAX query info for a table (check if it's DAX-backed) Related: tablecolumn (filter/sort/columns), datamodel (DAX measures, evaluate queries)

**Actions:** `list`, `create`, `rename`, `delete`, `read`, `resize`, `toggle-totals`, `set-column-total`, `append`, `get-data`, `set-style`, `add-to-data-model`, `create-from-dax`, `update-dax`, `get-dax`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet to create the table on (required for: create, create-from-dax) |
| `--table-name` | Name for the new table (must be unique in workbook) (required for: create, rename, delete, read, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-data-model, create-from-dax, update-dax, get-dax) |
| `--range` | Cell range address for the table (e.g., 'A1:D10') (required for: create) |
| `--has-headers` | True if first row contains column headers (default: true) |
| `--table-style` | Table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1'). Optional. (required for: set-style) |
| `--new-name` | New name for the table (must be unique in workbook) (required for: rename) |
| `--new-range` | New range address (e.g., 'A1:F20') (required for: resize) |
| `--show-totals` | True to show totals row, false to hide (required for: toggle-totals) |
| `--column-name` | Name of the column to set total function on (required for: set-column-total) |
| `--total-function` | Totals function name: Sum, Count, Average, Min, Max, CountNums, StdDev, Var, None (required for: set-column-total) |
| `--rows` | 2D array of row data to append - column order must match table columns. Optional if rowsFile is provided. (JSON format) |
| `--rows-file` | Path to a JSON or CSV file containing the rows to append. JSON: 2D array. CSV: rows/columns. Alternative to inline rows parameter |
| `--visible-only` | True to return only visible (non-filtered) rows; false for all rows (default: false) |
| `--strip-bracket-column-names` | When true, renames source table columns that contain literal bracket characters (removes brackets) beforeadding to the Data Model. This modifies the Excel table column headers in the worksheet |
| `--dax-query` | DAX EVALUATE query (e.g., 'EVALUATE Sales' or 'EVALUATE SUMMARIZE(...) ') (required for: create-from-dax, update-dax) |
| `--target-cell` | Target cell address for table placement (default: 'A1') |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### tablecolumn

Table column, filtering, and sorting operations for Excel Tables (ListObjects). Use table for table-level lifecycle and data operations. FILTERING: - 'apply-filter': Simple criteria filter (e.g., ">100", "=Active", "<>Closed") - 'apply-filter-values': Filter by exact values (provide list of values to include) - 'clear-filters': Remove all active filters - 'get-filters': See current filter state SORTING: - 'sort': Single column sort (ascending/descending) - 'sort-multi': Multi-column sort (provide list of {columnName, ascending} objects) COLUMN MANAGEMENT: - 'add-column'/'remove-column'/'rename-column': Modify table structure NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', '0%', 'yyyy-mm-dd')

**Actions:** `apply-filter`, `apply-filter-values`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `get-column-number-format`, `set-column-number-format`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--table-name` | Name of the Excel table (required) |
| `--column-name` | Name of the column to filter (required for: apply-filter, apply-filter-values, add-column, remove-column, sort, get-column-number-format, set-column-number-format) |
| `--criteria` | Filter criteria string (e.g., '>100', '=Active', '<>Closed') (required for: apply-filter) |
| `--values` | List of exact values to include in the filter (required for: apply-filter-values) (JSON format) |
| `--position` | 1-based column position (optional, defaults to end of table) |
| `--old-name` | Current column name (required for: rename-column) |
| `--new-name` | New column name (required for: rename-column) |
| `--region` | Table region: 'Data', 'Headers', 'Totals', or 'All' (required for: get-structured-reference) |
| `--ascending` | Sort order: true = ascending (A-Z, 0-9), false = descending (default: true) |
| `--sort-columns` | List of sort specifications: [{columnName: 'Col1', ascending: true}, ...] - applied in order (required for: sort-multi) (JSON format) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00', '0%', 'yyyy-mm-dd') (required for: set-column-number-format) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### vba

VBA module and procedure operations for macro-enabled workbooks (.xlsm). PREREQUISITES: - Workbook must be macro-enabled (.xlsm) - VBA trust must be enabled manually in Excel for project access SCOPE: - List and view existing VBA components and their procedures - Import creates new standard modules from inline code or a file - Update/delete works on existing VBA components by name - Run executes a procedure by name RUN: procedureName format is 'Module.Procedure' (e.g., 'Module1.MySub'). ExcelMcp does not configure VBA trust settings for you

**Actions:** `list`, `view`, `import`, `update`, `run`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--module-name` | Name of the VBA module (required for: view, import, update, delete) |
| `--vba-code` | VBA code to import (required for: import, update) |
| `--vba-code-file` | Path to file containing vbaCode |
| `--procedure-name` | Name of the procedure to run (for example "Module1.MySub") (required for: run) |
| `--timeout` | Optional timeout for execution Accepts seconds (e.g. 600) or TimeSpan format (e.g. 00:10:00) |
| `--parameters` | Optional parameters to pass to the procedure (required for: run) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### window

Control Excel window visibility, position, state, status bar, and worksheet-specific views. Use to show/hide Excel, bring it to front, reposition, maximize/minimize, freeze panes, split panes, set zoom, and control gridlines, headings, and outline symbols. Set status bar text to give users real-time feedback during operations. VISIBILITY: 'show' makes Excel visible AND brings to front. 'hide' hides Excel. Visibility changes are reflected in session metadata (session list shows updated state). WINDOW STATE values: 'normal', 'minimized', 'maximized'. ARRANGE presets: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen'. STATUS BAR: 'set-status-bar' displays text in Excel's status bar. 'clear-status-bar' restores default. WORKSHEET VIEW: View properties belong to a workbook window and apply to the named active worksheet. 'freeze-panes' uses row and column counts above/left of the pane boundary. 'set-split' creates movable panes and disables frozen panes. Zoom must be between 10 and 400 percent

**Actions:** `show`, `hide`, `bring-to-front`, `get-info`, `set-state`, `set-position`, `arrange`, `set-status-bar`, `clear-status-bar`, `get-view`, `freeze-panes`, `unfreeze-panes`, `set-split`, `set-zoom`, `set-display-options`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--window-state` | Window state: 'normal', 'minimized', or 'maximized' (required for: set-state) |
| `--left` | Window left position in points |
| `--top` | Window top position in points |
| `--width` | Window width in points |
| `--height` | Window height in points |
| `--preset` | Preset name: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen' (required for: arrange) |
| `--text` | Status bar text to display (e.g. "Building PivotTable from Sales data...") (required for: set-status-bar) |
| `--sheet` | Worksheet whose view should be inspected (required for: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, set-display-options) |
| `--frozen-rows` | Number of rows to freeze from the top (0-1,048,575) |
| `--frozen-columns` | Number of columns to freeze from the left (0-16,383) |
| `--split-rows` | Number of rows above the horizontal split (0-1,048,575) |
| `--split-columns` | Number of columns left of the vertical split (0-16,383) |
| `--zoom` | Zoom percentage from 10 through 400 (required for: set-zoom) |
| `--show-gridlines` | Whether to display cell gridlines |
| `--show-headings` | Whether to display row and column headings |
| `--show-outline-symbols` | Whether to display outline level symbols |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### workbook

Manage workbook metadata, document properties, Save As/copy operations, fixed-format exports, and external Excel links. SAVE-AS formats: auto, xlsx, xlsm, xlsb, xls. The active session follows the new workbook path. FIXED FORMAT: PDF or XPS with standard or minimum quality. DOCUMENT PROPERTIES: built-in properties can be read/updated; custom properties can be created, updated, and deleted. EXTERNAL LINKS: discovers, updates, or permanently breaks Excel workbook links. Printing and print preview are intentionally excluded because default-printer output and modal preview are unsafe for unattended automation

**Actions:** `get-info`, `list-document-properties`, `get-document-property`, `set-document-property`, `delete-document-property`, `save-as`, `save-copy-as`, `export-fixed-format`, `list-external-links`, `update-external-link`, `break-external-link`, `set-protection`, `get-protection`, `set-view-options`, `get-view-options`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--include-built-in` | IncludeBuil tIn |
| `--include-custom` | IncludeCust om |
| `--property-name` | Document property name (required for: get-document-property, set-document-property, delete-document-property) |
| `--scope` | Property collection: built-in or custom |
| `--value` | String value to store (required for: set-document-property) |
| `--target-path` | Absolute output path in an existing directory (required for: save-as, save-copy-as, export-fixed-format) |
| `--format` | Output format: auto, xlsx, xlsm, xlsb, or xls |
| `--overwrite` | Whether an existing output file may be replaced |
| `--format-type` | FormatType |
| `--quality` | Quality |
| `--include-document-properties` | IncludeDocu mentPropert ies |
| `--ignore-print-areas` | IgnorePrint Areas |
| `--from-page` | FromPage |
| `--to-page` | ToPage |
| `--open-after-publish` | OpenAfterPu blish |
| `--link-source` | (required for: update-external-link, break-external-link) |
| `--is-protected` | (required for: set-protection) |
| `--password` | Password |
| `--display-gridlines` | DisplayGrid lines |
| `--display-headings` | DisplayHead ings |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### worksheetstyle

Worksheet styling, visibility, protection, grouping, and outline operations. Use sheet for lifecycle operations (create, rename, copy, delete, move). TAB COLORS: Use RGB values (0-255 each) to set custom tab colors for visual organization. VISIBILITY LEVELS: - 'visible': Normal visible sheet - 'hidden': Hidden but accessible via Format > Sheet > Unhide - 'veryhidden': Only accessible via VBA (protection against casual unhiding) PROTECTION: Protect a worksheet to lock its contents and structure, or unprotect it. OUTLINES: Group or ungroup row/column ranges, configure summary positions, show a specific row/column outline level, inspect grouping state, or clear all groups

**Actions:** `set-tab-color`, `get-tab-color`, `clear-tab-color`, `set-protection`, `get-protection`, `set-comment`, `get-comment`, `clear-comment`, `add-image`, `get-image-count`, `add-shape`, `get-shape-count`, `set-page-setup`, `get-page-setup`, `set-visibility`, `get-visibility`, `show`, `hide`, `very-hide`, `group`, `ungroup`, `get-outline-info`, `set-outline-settings`, `show-outline-levels`, `clear-outline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet to color (required) |
| `--red` | Red color component (0-255) (required for: set-tab-color) |
| `--green` | Green color component (0-255) (required for: set-tab-color) |
| `--blue` | Blue color component (0-255) (required for: set-tab-color) |
| `--is-protected` | Whether the worksheet should be protected (required for: set-protection) |
| `--password` | Optional password for protecting/unprotecting the sheet |
| `--cell-address` | Cell address such as A1 (required for: set-comment, get-comment, clear-comment, add-image, add-shape) |
| `--text` | Cell note text to set (required for: set-comment) |
| `--image-path` | Absolute path to the image file on disk (required for: add-image) |
| `--orientation` | Page orientation: 'portrait' or 'landscape' (required for: set-page-setup) |
| `--fit-to-pages-wide` | Number of pages wide to fit the printout to |
| `--fit-to-pages-tall` | Number of pages tall to fit the printout to |
| `--center-horizontally` | Whether to center the printout horizontally on the page |
| `--center-vertically` | Whether to center the printout vertically on the page |
| `--visibility` | Visibility level: 'visible', 'hidden', or 'veryhidden' (required for: set-visibility) |
| `--range` | Row or column range to group (required for: group, ungroup, get-outline-info) |
| `--axis` | Grouping axis: Rows or Columns (required for: group, ungroup, get-outline-info) |
| `--summary-row` | Summary row position: above or below |
| `--summary-column` | Summary column position: left or right |
| `--automatic-styles` | Whether Excel applies automatic outline styles |
| `--row-levels` | Optional row outline level to display |
| `--column-levels` | Optional column outline level to display |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### xmlmap

Manage workbook XML maps and exchange XML data without interactive dialogs. SECURITY: Schemas and XML data are parsed from supplied content with DTD processing disabled. XSD import/include/redefine dependencies and XML schema-location attributes are rejected to prevent implicit network or file access. IMPORT MODES: Provide mapName to import into existing mapped cells. Omit mapName and provide sheetName plus startCell to let Excel create a map and XML table at that destination

**Actions:** `list`, `add`, `map-range`, `import-xml`, `export-xml`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--schema` | Inline XSD schema content (required for: add) |
| `--schema-file` | Path to file containing schema |
| `--root-element-name` | Optional root element when the schema has multiple roots |
| `--map-name` | Optional name to assign to the created map (required for: map-range, export-xml, delete) |
| `--sheet` | Worksheet containing the target range (required for: map-range) |
| `--range` | Target cell or single-column range (required for: map-range) |
| `--xpath` | XPath to map (required for: map-range) |
| `--selection-namespace` | Optional namespace declarations used by prefixed XPath expressions |
| `--repeating` | Whether to create a repeating XML list mapping |
| `--xml-data` | Inline XML data (required for: import-xml) |
| `--xml-data-file` | Path to file containing xmlData |
| `--start-cell` | Top-left destination cell for automatic mapping |
| `--overwrite` | Whether imported XML may overwrite mapped cells |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

## Common Pitfalls

- `--values-file` requires an existing JSON or CSV file; use `--values` for inline JSON.
- `--timeout` must be greater than zero; omit it to use the command default.
- `pythoninexcel get-result --max-wait-seconds` must be at least 1 and shorter than the session operation timeout.
- `--values` and list parameters use JSON arrays; range values use a two-dimensional array.
- Power Query operations may take 30 seconds or longer; use a generous positive timeout.

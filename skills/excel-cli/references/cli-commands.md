# CLI Command Reference

> Auto-generated from the built `excelcli` runtime. Use these exact command and parameter names.

### analysis

Excel what-if analysis with Goal Seek, scenarios, scenario reports, and one- or two-variable data tables

**Actions:** `goal-seek`, `list-scenarios`, `create-scenario`, `update-scenario`, `show-scenario`, `delete-scenario`, `create-scenario-summary`, `create-data-table`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | (required) |
| `--formula-cell` | (required for: goal-seek) (valid for: goal-seek) |
| `--goal` | (required for: goal-seek) (valid for: goal-seek) |
| `--changing-cell` | (required for: goal-seek) (valid for: goal-seek) |
| `--scenario-name` | (required for: create-scenario, update-scenario, show-scenario, delete-scenario) (valid for: create-scenario, update-scenario, show-scenario, delete-scenario) |
| `--changing-cells` | (required for: create-scenario, update-scenario) (valid for: create-scenario, update-scenario) |
| `--values` | (required for: create-scenario, update-scenario) (valid for: create-scenario, update-scenario) (JSON format) |
| `--comment` | (valid for: create-scenario) |
| `--locked` | (valid for: create-scenario) |
| `--hidden` | (valid for: create-scenario) |
| `--report-type` | (valid for: create-scenario-summary) |
| `--result-cells` | (valid for: create-scenario-summary) |
| `--table-range` | (required for: create-data-table) (valid for: create-data-table) |
| `--row-input-cell` | (valid for: create-data-table) |
| `--column-input-cell` | (valid for: create-data-table) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### batch

Execute multiple commands from a JSON file or stdin. Outputs NDJSON (one result per line)

| Parameter | Description |
|-----------|-------------|
| `--input` | JSON file with command array. Use '-' for stdin (NDJSON, one command per line). If omitted, reads from stdin |
| `--session` | Default session ID for all commands. Overridden by per-command sessionId. Auto-captured from session.open/create if not set |
| `--stop-on-error` | Stop execution on first error (default: continue all commands) |

### calculationmode

Control Excel recalculation (automatic vs manual)

**Actions:** `get-mode`, `set-mode`, `calculate`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--mode` | Target calculation mode (required for: set-mode) (valid for: set-mode) |
| `--scope` | Scope: Workbook, Sheet, or Range (required for: calculate) (valid for: calculate) |
| `--sheet` | Sheet name (required for Sheet/Range scope) (valid for: calculate) |
| `--range` | Range address (required for Range scope) (valid for: calculate) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### chart

Chart lifecycle - create, read, move, and delete embedded charts

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-pivottable`, `delete`, `move`, `fit-to-range`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--chart-name` | Name of the chart (or shape name) (required for: read, delete, move, fit-to-range) (valid for: read, create-from-range, create-from-table, create-from-pivottable, delete, move, fit-to-range) |
| `--sheet` | Target worksheet name (required for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) (valid for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) |
| `--source-range-address` | Data range for the chart (e.g., A1:D10) (required for: create-from-range) (valid for: create-from-range) |
| `--chart-type` | Type of chart to create (required for: create-from-range, create-from-table, create-from-pivottable) (valid for: create-from-range, create-from-table, create-from-pivottable) |
| `--left` | Left position in points from worksheet edge (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--top` | Top position in points from worksheet edge (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--width` | Chart width in points (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--height` | Chart height in points (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--target-range` | Cell range to position chart within (e.g., 'F2:K15'). PREFERRED over left/top. When set, left/top are ignored. (valid for: create-from-range, create-from-table, create-from-pivottable) |
| `--table-name` | Name of the Excel Table (required for: create-from-table) (valid for: create-from-table) |
| `--pivot-table-name` | Name of the source PivotTable (required for: create-from-pivottable) (valid for: create-from-pivottable) |
| `--range` | Range to fit the chart to (e.g., A1:D10) (required for: fit-to-range) (valid for: fit-to-range) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### chartconfig

Chart configuration - data source, series, type, title, axis labels, legend, and styling

**Actions:** `set-source-range`, `add-series`, `remove-series`, `set-chart-type`, `set-title`, `set-axis-title`, `get-axis-number-format`, `set-axis-number-format`, `show-legend`, `set-style`, `set-placement`, `set-data-labels`, `get-axis-scale`, `set-axis-scale`, `get-gridlines`, `set-gridlines`, `set-series-format`, `set-series-chart-type`, `get-plot-options`, `set-plot-options`, `set-area-format`, `list-trendlines`, `add-trendline`, `delete-trendline`, `set-trendline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--chart-name` | Name of the chart (required) |
| `--source-range` | New data source range (e.g., Sheet1!A1:D10) (required for: set-source-range) (valid for: set-source-range) |
| `--series-name` | Display name for the series (required for: add-series) (valid for: add-series) |
| `--values-range` | Range containing series values (e.g., B2:B10) (required for: add-series) (valid for: add-series) |
| `--category-range` | Optional range for category labels (e.g., A2:A10) (valid for: add-series) |
| `--series-index` | 1-based index of the series to remove (required for: remove-series, set-series-format, set-series-chart-type, list-trendlines, add-trendline, delete-trendline, set-trendline) (valid for: remove-series, set-data-labels, set-series-format, set-series-chart-type, list-trendlines, add-trendline, delete-trendline, set-trendline) |
| `--chart-type` | New chart type to apply (required for: set-chart-type, set-series-chart-type) (valid for: set-chart-type, set-series-chart-type) |
| `--title` | Title text to display (required for: set-title, set-axis-title) (valid for: set-title, set-axis-title) |
| `--axis` | Which axis to set title for (Category, Value, SeriesAxis) (required for: set-axis-title, get-axis-number-format, set-axis-number-format, get-axis-scale, set-axis-scale, set-gridlines) (valid for: set-axis-title, get-axis-number-format, set-axis-number-format, get-axis-scale, set-axis-scale, set-gridlines) |
| `--number-format` | Excel number format code (e.g., "$#,##0", "0.00%") (required for: set-axis-number-format) (valid for: set-axis-number-format) |
| `--visible` | True to show legend, false to hide (required for: show-legend) (valid for: show-legend) |
| `--legend-position` | Optional position for the legend (valid for: show-legend) |
| `--style-id` | Excel chart style ID (1-48 for most chart types) (required for: set-style) (valid for: set-style) |
| `--placement` | Placement mode: 1=MoveAndSize, 2=Move, 3=FreeFloating (required for: set-placement) (valid for: set-placement) |
| `--print-object` | Whether the embedded chart prints with the worksheet (valid for: set-placement) |
| `--locked` | Whether the embedded chart is locked when the worksheet is protected (valid for: set-placement) |
| `--rounded-corners` | Whether the embedded chart uses rounded corners (valid for: set-placement) |
| `--show-value` | Show data values on labels (valid for: set-data-labels) |
| `--show-percentage` | Show percentage values. Only meaningful for pie and doughnut chart types; setting to true on other chart types has no visual effect. (valid for: set-data-labels) |
| `--show-series-name` | Show series name on labels (valid for: set-data-labels) |
| `--show-category-name` | Show category name on labels (valid for: set-data-labels) |
| `--show-bubble-size` | Show bubble size (bubble charts) (valid for: set-data-labels) |
| `--separator` | Separator string between label components (valid for: set-data-labels) |
| `--label-position` | Position of data labels relative to data points (valid for: set-data-labels) |
| `--minimum-scale` | Minimum axis value (null for auto) (valid for: set-axis-scale) |
| `--maximum-scale` | Maximum axis value (null for auto) (valid for: set-axis-scale) |
| `--major-unit` | Major gridline interval (null for auto) (valid for: set-axis-scale) |
| `--minor-unit` | Minor gridline interval (null for auto) (valid for: set-axis-scale) |
| `--show-major` | Show major gridlines (null to keep current) (valid for: set-gridlines) |
| `--show-minor` | Show minor gridlines (null to keep current) (valid for: set-gridlines) |
| `--marker-style` | Marker shape style (valid for: set-series-format) |
| `--marker-size` | Marker size in points (2-72) (valid for: set-series-format) |
| `--marker-background-color` | Marker fill color (#RRGGBB) (valid for: set-series-format) |
| `--marker-foreground-color` | Marker border color (#RRGGBB) (valid for: set-series-format) |
| `--invert-if-negative` | Invert colors for negative values (valid for: set-series-format) |
| `--fill-color` | Series fill color as #RRGGBB (valid for: set-series-format, set-area-format) |
| `--fill-transparency` | Series fill transparency from 0 (opaque) to 1 (transparent) (valid for: set-series-format, set-area-format) |
| `--line-color` | Series line color as #RRGGBB (valid for: set-series-format, set-area-format) |
| `--line-weight` | Series line weight in points (valid for: set-series-format, set-area-format) |
| `--plot-by` | Interpret source rows or columns as data series (valid for: set-plot-options) |
| `--display-blanks-as` | How blank cells appear: gaps, zeroes, or interpolation (valid for: set-plot-options) |
| `--plot-visible-only` | True to omit hidden rows and columns (valid for: set-plot-options) |
| `--area` | Chart area or plot area (required for: set-area-format) (valid for: set-area-format) |
| `--trendline-type` | Type of trendline (Linear, Exponential, etc.) (required for: add-trendline) (valid for: add-trendline) |
| `--order` | Polynomial order (2-6, for Polynomial type) (valid for: add-trendline) |
| `--period` | Moving average period (for MovingAverage type) (valid for: add-trendline) |
| `--forward` | Periods to extend forward (valid for: add-trendline, set-trendline) |
| `--backward` | Periods to extend backward (valid for: add-trendline, set-trendline) |
| `--intercept` | Force trendline through specific Y-intercept (valid for: add-trendline, set-trendline) |
| `--display-equation` | Display trendline equation on chart (valid for: add-trendline, set-trendline) |
| `--display-r-squared` | Display R-squared value on chart (valid for: add-trendline, set-trendline) |
| `--name` | Custom name for the trendline (valid for: add-trendline, set-trendline) |
| `--trendline-index` | 1-based index of the trendline to delete (required for: delete-trendline, set-trendline) (valid for: delete-trendline, set-trendline) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### conditionalformat

Conditional formatting - visual rules based on cell values

**Actions:** `add-rule`, `clear-rules`, `list-rules`, `list-worksheet-rules`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Sheet name (empty for active sheet) (required) |
| `--range` | Range address (A1 notation or named range) (required for: add-rule, clear-rules, list-rules) (valid for: add-rule, clear-rules, list-rules) |
| `--rule-type` | Rule type: cellValue (or cell-value), expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage. Both camelCase and kebab-case accepted. (required for: add-rule) (valid for: add-rule) |
| `--operator-type` | Required for cellValue rules. XlFormatConditionOpe rator: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween (valid for: add-rule) |
| `--formula1` | Required for cellValue and expression rules. First formula/value for condition (valid for: add-rule) |
| `--formula2` | Required for between/notBetween cellValue rules. Second formula/value (valid for: add-rule) |
| `--interior-color` | Fill color (#RRGGBB or color index) (valid for: add-rule) |
| `--interior-pattern` | Interior pattern (1=Solid, -4142=None, 9=Gray50, etc.) (valid for: add-rule) |
| `--font-color` | Font color (#RRGGBB or color index) (valid for: add-rule) |
| `--font-bold` | Bold font (valid for: add-rule) |
| `--font-italic` | Italic font (valid for: add-rule) |
| `--border-style` | Border style: none, continuous, dash, dot, etc. (valid for: add-rule) |
| `--border-color` | Border color (#RRGGBB or color index) (valid for: add-rule) |
| `--color-scale-min-type` | colorScale minimum stop type: minimum, number, percent, percentile, formula (valid for: add-rule) |
| `--color-scale-min-value` | colorScale minimum stop value (for number/percent/perce ntile/formula) (valid for: add-rule) |
| `--color-scale-min-color` | colorScale minimum stop color (#RRGGBB) (valid for: add-rule) |
| `--color-scale-mid-type` | colorScale midpoint stop type (supply to create a 3-color scale) (valid for: add-rule) |
| `--color-scale-mid-value` | colorScale midpoint stop value (valid for: add-rule) |
| `--color-scale-mid-color` | colorScale midpoint stop color (#RRGGBB) (valid for: add-rule) |
| `--color-scale-max-type` | colorScale maximum stop type: maximum, number, percent, percentile, formula (valid for: add-rule) |
| `--color-scale-max-value` | colorScale maximum stop value (valid for: add-rule) |
| `--color-scale-max-color` | colorScale maximum stop color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-color` | dataBar fill color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-negative-color` | dataBar negative-value bar color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-direction` | dataBar fill direction: context, leftToRight, rightToLeft (valid for: add-rule) |
| `--data-bar-show-value` | dataBar show the cell value alongside the bar (valid for: add-rule) |
| `--data-bar-min-type` | dataBar minimum point type: automaticMinimum, minimum, number, percent, percentile, formula (valid for: add-rule) |
| `--data-bar-min-value` | dataBar minimum point value (valid for: add-rule) |
| `--data-bar-max-type` | dataBar maximum point type: automaticMaximum, maximum, number, percent, percentile, formula (valid for: add-rule) |
| `--data-bar-max-value` | dataBar maximum point value (valid for: add-rule) |
| `--icon-set-id` | iconSet id: 3Arrows, 3TrafficLights1, 4Ratings, 5Quarters, etc. (valid for: add-rule) |
| `--icon-set-reverse` | iconSet reverse icon order (valid for: add-rule) |
| `--icon-set-show-icon-only` | iconSet show only the icon (hide the value) (valid for: add-rule) |
| `--icon-threshold1-type` | iconSet threshold 1 type: percent, number, percentile, formula (valid for: add-rule) |
| `--icon-threshold1-value` | iconSet threshold 1 value (valid for: add-rule) |
| `--icon-threshold2-type` | iconSet threshold 2 type (valid for: add-rule) |
| `--icon-threshold2-value` | iconSet threshold 2 value (valid for: add-rule) |
| `--icon-threshold3-type` | iconSet threshold 3 type (valid for: add-rule) |
| `--icon-threshold3-value` | iconSet threshold 3 value (valid for: add-rule) |
| `--icon-threshold4-type` | iconSet threshold 4 type (valid for: add-rule) |
| `--icon-threshold4-value` | iconSet threshold 4 value (valid for: add-rule) |
| `--rank` | top10 rank (number of values, or percent when top10Percent is true) (valid for: add-rule) |
| `--top10-percent` | top10 treat rank as a percentage (valid for: add-rule) |
| `--top-bottom` | top10 direction: top or bottom (valid for: add-rule) |
| `--above-below` | aboveAverage selector: aboveAverage, belowAverage, aboveStdDev, belowStdDev, equalAboveAverage, equalBelowAverage (valid for: add-rule) |
| `--date-period` | timePeriod period: today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth (valid for: add-rule) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### connection

Data connections (OLEDB, ODBC, ODC import)

**Actions:** `list`, `view`, `create`, `refresh`, `get-refresh-status`, `cancel-refresh`, `delete`, `load-to`, `get-properties`, `set-properties`, `test`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--connection-name` | Name of the connection to view (required for: view, create, refresh, get-refresh-status, cancel-refresh, delete, load-to, get-properties, set-properties, test) (valid for: view, create, refresh, get-refresh-status, cancel-refresh, delete, load-to, get-properties, set-properties, test) |
| `--connection-string` | OLEDB or ODBC connection string (required for: create) (valid for: create, set-properties) |
| `--command-text` | SQL query or table name (valid for: create, set-properties) |
| `--description` | Optional description for the connection (valid for: create, set-properties) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--sheet` | Target worksheet name (required for: load-to) (valid for: load-to) |
| `--background-query` | Run query in background (null to keep current) (valid for: set-properties) |
| `--refresh-on-file-open` | Refresh when file opens (null to keep current) (valid for: set-properties) |
| `--save-password` | Save password in connection (null to keep current) (valid for: set-properties) |
| `--refresh-period` | Auto-refresh interval in minutes (null to keep current) (valid for: set-properties) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### datamodel

Data Model (Power Pivot) - DAX measures and table management

**Actions:** `list-tables`, `list-columns`, `read-table`, `read-info`, `read-connection`, `list-measures`, `read`, `delete-measure`, `delete-table`, `rename-table`, `refresh`, `create-measure`, `update-measure`, `evaluate`, `execute-dmv`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--table-name` | Name of the table to list columns from (required for: list-columns, read-table, delete-table, create-measure) (valid for: list-columns, read-table, list-measures, delete-table, refresh, create-measure) |
| `--measure-name` | Name of the measure to get (required for: read, delete-measure, create-measure, update-measure) (valid for: read, delete-measure, create-measure, update-measure) |
| `--old-name` | Current name of the table (required for: rename-table) (valid for: rename-table) |
| `--new-name` | New name for the table (required for: rename-table) (valid for: rename-table) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--dax-formula` | DAX formula. Public callers must supply either inline daxFormula or a readable daxFormulaFile, not both. (required for: create-measure) (valid for: create-measure, update-measure) |
| `--dax-formula-file` | Path to a readable file containing daxFormula; use instead of inline daxFormula, not together (valid for: create-measure, update-measure) |
| `--format-type` | Optional format type: General, Currency, Decimal, Percentage, or WholeNumber (case-insensitive). Null or empty defaults to General on create and keeps the existing format on update. (valid for: create-measure, update-measure) |
| `--description` | Optional: Description of the measure (valid for: create-measure, update-measure) |
| `--format-dax` | Whether to send the DAX formula to the remote daxformatter.com service before saving. Defaults to false to preserve privacy. (valid for: create-measure, update-measure) |
| `--dax-query` | DAX EVALUATE query. Public callers must supply either inline daxQuery or a readable daxQueryFile, not both. (required for: evaluate) (valid for: evaluate) |
| `--dax-query-file` | Path to a readable file containing daxQuery; use instead of inline daxQuery, not together (valid for: evaluate) |
| `--dmv-query` | DMV query in SQL-like syntax. Public callers must supply either inline dmvQuery or a readable dmvQueryFile, not both. (required for: execute-dmv) (valid for: execute-dmv) |
| `--dmv-query-file` | Path to a readable file containing dmvQuery; use instead of inline dmvQuery, not together (valid for: execute-dmv) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### datamodelrelationship

Data Model relationships - link tables for cross-table DAX calculations

**Actions:** `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--from-table` | Source table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--from-column` | Source column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-table` | Target table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-column` | Target column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--active` | Whether the relationship should be active (default: true) (required for: update-relationship) (valid for: create-relationship, update-relationship) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### diag

Diagnostic commands for testing CLI/MCP infrastructure without Excel

**Actions:** `ping`, `echo`, `validate-params`

| Parameter | Description |
|-----------|-------------|
| `--message` | The message to echo back (required) (required for: echo) (valid for: echo) |
| `--tag` | Optional tag to include in the response (valid for: echo) |
| `--name` | Required name parameter (required for: validate-params) (valid for: validate-params) |
| `--count` | Required integer parameter (required for: validate-params) (valid for: validate-params) |
| `--label` | Optional label parameter (valid for: validate-params) |
| `--verbose` | Optional boolean flag (default: false) (valid for: validate-params) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### drawing

Worksheet drawing objects and sparklines

**Actions:** `list-objects`, `get-object`, `add-image`, `add-shape`, `add-text-box`, `add-connector`, `add-form-control`, `update-object`, `delete-object`, `list-sparklines`, `get-sparkline`, `add-sparkline`, `update-sparkline`, `delete-sparkline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | (required) |
| `--object-name` | (required for: get-object, update-object, delete-object) (valid for: get-object, update-object, delete-object) |
| `--image-path` | (required for: add-image) (valid for: add-image) |
| `--name` | (valid for: add-image, add-shape, add-text-box, add-connector, add-form-control) |
| `--left` | (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--top` | (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--width` | (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--height` | (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--lock-aspect-ratio` | (valid for: add-image) |
| `--shape-type` | (valid for: add-shape) |
| `--text` | (required for: add-text-box) (valid for: add-shape, add-text-box, add-form-control, update-object) |
| `--fill-color` | (valid for: add-shape, add-text-box, update-object) |
| `--line-color` | (valid for: add-shape, add-text-box, add-connector, update-object, add-sparkline, update-sparkline) |
| `--line-weight` | (valid for: add-shape, add-connector, update-object) |
| `--font-size` | (valid for: add-text-box, update-object) |
| `--font-color` | (valid for: add-text-box, update-object) |
| `--connector-type` | (valid for: add-connector) |
| `--begin-x` | (valid for: add-connector) |
| `--begin-y` | (valid for: add-connector) |
| `--end-x` | (valid for: add-connector) |
| `--end-y` | (valid for: add-connector) |
| `--control-type` | (valid for: add-form-control) |
| `--linked-cell` | (valid for: add-form-control, update-object) |
| `--input-range` | (valid for: add-form-control, update-object) |
| `--new-name` | (valid for: update-object) |
| `--rotation` | (valid for: update-object) |
| `--visible` | (valid for: update-object) |
| `--locked` | (valid for: update-object) |
| `--placement` | (valid for: update-object) |
| `--alternative-text` | (valid for: update-object) |
| `--location-range` | (required for: get-sparkline, add-sparkline, update-sparkline, delete-sparkline) (valid for: get-sparkline, add-sparkline, update-sparkline, delete-sparkline) |
| `--source-range` | (required for: add-sparkline) (valid for: add-sparkline, update-sparkline) |
| `--sparkline-type` | (valid for: add-sparkline, update-sparkline) |
| `--show-markers` | (valid for: add-sparkline, update-sparkline) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### namedrange

Named ranges for formulas/parameters

**Actions:** `list`, `write`, `read`, `update`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--name` | Name of the named range (required for: write, read, update, create, delete) (valid for: write, read, update, create, delete) |
| `--value` | Value to set (required for: write) (valid for: write) |
| `--reference` | New cell reference (e.g., Sheet1!$A$1:$B$10) (required for: update, create) (valid for: update, create) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pivottable

PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-datamodel`, `delete`, `refresh`, `get-cache-options`, `set-cache-options`, `drill-through`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh, get-cache-options, set-cache-options, drill-through) (valid for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh, get-cache-options, set-cache-options, drill-through) |
| `--source-sheet` | Source worksheet name (required for: create-from-range) (valid for: create-from-range) |
| `--source-range` | Source range address (e.g., "A1:F100") (required for: create-from-range) (valid for: create-from-range) |
| `--destination-sheet` | Destination worksheet name (required for: create-from-range, create-from-table, create-from-datamodel) (valid for: create-from-range, create-from-table, create-from-datamodel) |
| `--destination-cell` | Destination cell address (e.g., "A1") (required for: create-from-range, create-from-table, create-from-datamodel) (valid for: create-from-range, create-from-table, create-from-datamodel) |
| `--table-name` | Name of the Excel Table (required for: create-from-table, create-from-datamodel) (valid for: create-from-table, create-from-datamodel) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--enable-refresh` | Allow the PivotCache to refresh (valid for: set-cache-options) |
| `--refresh-on-file-open` | Refresh the PivotCache when the workbook opens (valid for: set-cache-options) |
| `--missing-items-limit` | How many deleted source items Excel retains in the cache (valid for: set-cache-options) |
| `--optimize-cache` | Optimize the cache when it is constructed (valid for: set-cache-options) |
| `--save-source-data` | Save source records with the PivotTable (valid for: set-cache-options) |
| `--cell-address` | Value-cell address on the PivotTable worksheet (for example, G4) (required for: drill-through) (valid for: drill-through) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pivottablecalc

PivotTable calculated fields/members, layout configuration, and data extraction

**Actions:** `get-data`, `create-calculated-field`, `list-calculated-fields`, `delete-calculated-field`, `list-calculated-members`, `create-calculated-member`, `delete-calculated-member`, `set-layout`, `set-subtotals`, `set-grand-totals`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required) |
| `--field-name` | Name for the calculated field (required for: create-calculated-field, delete-calculated-field, set-subtotals) (valid for: create-calculated-field, delete-calculated-field, set-subtotals) |
| `--formula` | Formula using field references (e.g., "=Revenue-Cost") (required for: create-calculated-field, create-calculated-member) (valid for: create-calculated-field, create-calculated-member) |
| `--member-name` | Name for the calculated member (MDX naming format) (required for: create-calculated-member, delete-calculated-member) (valid for: create-calculated-member, delete-calculated-member) |
| `--type` | Type of calculated member (Member, Set, or Measure) (valid for: create-calculated-member) |
| `--solve-order` | Solve order for calculation precedence (default: 0) (valid for: create-calculated-member) |
| `--display-folder` | Display folder path for organizing measures (optional) (valid for: create-calculated-member) |
| `--number-format` | Number format code for the calculated member (optional) (valid for: create-calculated-member) |
| `--row-layout` | Layout form: 0=Compact, 1=Tabular, 2=Outline (required for: set-layout) (valid for: set-layout) |
| `--show-subtotals` | True to show automatic subtotals, false to hide (required for: set-subtotals) (valid for: set-subtotals) |
| `--show-row-grand-totals` | Show row grand totals (bottom summary row) (required for: set-grand-totals) (valid for: set-grand-totals) |
| `--show-column-grand-totals` | Show column grand totals (right summary column) (required for: set-grand-totals) (valid for: set-grand-totals) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pivottablefield

PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping

**Actions:** `list-fields`, `add-row-field`, `add-column-field`, `add-value-field`, `add-filter-field`, `remove-field`, `set-field-function`, `set-field-name`, `set-field-format`, `set-field-filter`, `sort-field`, `group-by-date`, `group-by-numeric`, `group-items`, `ungroup-field`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required) |
| `--field-name` | Name of the field to add (required for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, group-by-date, group-by-numeric, group-items) (valid for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, group-by-date, group-by-numeric, group-items) |
| `--position` | Optional position in row area (1-based) (valid for: add-row-field, add-column-field) |
| `--aggregation-function` | Aggregation function (for Regular and OLAP auto-create mode) (required for: set-field-function) (valid for: add-value-field, set-field-function) |
| `--custom-name` | Optional custom name for the field/measure (required for: set-field-name) (valid for: add-value-field, set-field-name) |
| `--number-format` | Number format string (required for: set-field-format) (valid for: set-field-format) |
| `--selected-values` | Values to show (others will be hidden) (required for: set-field-filter) (valid for: set-field-filter) (JSON format) |
| `--direction` | Sort direction (valid for: sort-field) |
| `--interval` | Grouping interval (Months, Quarters, Years) (required for: group-by-date) (valid for: group-by-date) |
| `--start` | Starting value (null = use field minimum) (valid for: group-by-numeric) |
| `--end-value` | Ending value (null = use field maximum) (valid for: group-by-numeric) |
| `--interval-size` | Size of each group (e.g., 100 for groups of 100) (required for: group-by-numeric) (valid for: group-by-numeric) |
| `--item-names` | Exact item captions to group (required for: group-items) (valid for: group-items) (JSON format) |
| `--group-name` | Caption for the new group (required for: group-items) (valid for: group-items) |
| `--grouped-field-name` | Generated grouped field name (required for: ungroup-field) (valid for: ungroup-field) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### powerquery

Power Query M code and data loading

**Actions:** `list`, `view`, `refresh`, `get-load-config`, `delete`, `create`, `update`, `load-to`, `refresh-all`, `rename`, `unload`, `evaluate`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--query-name` | Name of the query to view (required for: view, refresh, get-load-config, delete, create, update, load-to, unload) (valid for: view, refresh, get-load-config, delete, create, update, load-to, unload) |
| `--timeout` | Public input is whole seconds from 0 through 2147483. Omitted or 0 uses the 30-minute data-operation default. (valid for: refresh, refresh-all) |
| `--m-code` | Raw M code. Public callers must supply either inline mCode or a readable mCodeFile, not both. (required for: create, update, evaluate) (valid for: create, update, evaluate) |
| `--m-code-file` | Path to a readable file containing mCode; use instead of inline mCode, not together (valid for: create, update, evaluate) |
| `--load-destination` | Load destination mode (required for: load-to) (valid for: create, load-to) |
| `--target-sheet` | Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted) (valid for: create, load-to) |
| `--target-cell-address` | Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data. (valid for: create, load-to) |
| `--format-m-code` | Whether to send M code to the remote powerqueryformatter.com service before saving. Defaults to false to preserve privacy. (valid for: create, update) |
| `--refresh` | Whether to refresh data after update (default: true) (valid for: update) |
| `--old-name` | Current name of the query (required for: rename) (valid for: rename) |
| `--new-name` | New name for the query (required for: rename) (valid for: rename) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### pythoninexcel

Microsoft 365 "Python in Excel" (=PY()) formulas

**Actions:** `set-formula`, `get-result`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet name (required) |
| `--range` | Target cell address (e.g. "D1") (required) |
| `--code` | Python source code (e.g. "xl('A1:A6').sum()") (required for: set-formula) (valid for: set-formula) |
| `--return-type` | 0 = Excel Value (default), 1 = Python Object (valid for: set-formula) |
| `--max-wait-seconds` | Maximum seconds to poll for the cloud result before giving up (default: 30). Must be shorter than the session operation timeout. Returns as soon as the result is ready. (valid for: get-result) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### querytable

Worksheet QueryTable lifecycle and configuration for local COM text, CSV, and legacy web imports

**Actions:** `list`, `view`, `create-text`, `create-web`, `set-properties`, `refresh`, `get-refresh-status`, `cancel-refresh`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | (required for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) (valid for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) |
| `--query-table-name` | (required for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) (valid for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) |
| `--source-path` | (required for: create-text) (valid for: create-text) |
| `--destination-address` | (required for: create-text, create-web) (valid for: create-text, create-web) |
| `--delimiter` | (valid for: create-text) |
| `--text-qualifier` | (valid for: create-text) |
| `--encoding` | (valid for: create-text) |
| `--has-headers` | (valid for: create-text) |
| `--url` | (required for: create-web) (valid for: create-web) |
| `--selection-type` | (valid for: create-web) |
| `--web-tables` | (valid for: create-web) |
| `--formatting` | (valid for: create-web) |
| `--background-query` | (valid for: set-properties) |
| `--refresh-on-file-open` | (valid for: set-properties) |
| `--refresh-period` | (valid for: set-properties) |
| `--adjust-column-width` | (valid for: set-properties) |
| `--preserve-formatting` | (valid for: set-properties) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### range

Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions

**Actions:** `get-values`, `set-values`, `get-formulas`, `set-formulas`, `validate-formulas`, `clear-all`, `clear-contents`, `clear-formats`, `copy`, `copy-values`, `copy-formulas`, `get-number-formats`, `set-number-format`, `set-number-formats`, `get-used-range`, `get-current-region`, `get-info`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only (required for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info) (valid for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info) |
| `--range` | Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name (e.g., 'SalesData') (required for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info) (valid for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info) |
| `--values` | 2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Optional if valuesFile is provided. (valid for: set-values) (JSON format) |
| `--values-file` | Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter. (valid for: set-values) |
| `--formulas` | 2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Optional if formulasFile is provided. (valid for: set-formulas, validate-formulas) (JSON format) |
| `--formulas-file` | Path to a JSON file containing the formulas as a 2D array. Alternative to inline formulas parameter. (valid for: set-formulas, validate-formulas) |
| `--source-sheet` | Source worksheet name for copy operations (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--source-range` | Source range address for copy operations (e.g., 'A1:D10') (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--target-sheet` | Target worksheet name for copy operations (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--target-range` | Target range address - can be single cell for paste destination (e.g., 'A1') (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00' for numbers, 'mm/dd/yyyy' for dates, '0.00%' for percentages, 'General' for default, '@' for text) (required for: set-number-format) (valid for: set-number-format) |
| `--formats` | 2D array of format codes - same dimensions as target range (e.g., [['#,##0.00', '0.00%'], ['mm/dd/yyyy', 'General']]). Optional if formatsFile is provided. (valid for: set-number-formats) (JSON format) |
| `--formats-file` | Path to a JSON file containing 2D array of format codes. Alternative to inline formats parameter. (valid for: set-number-formats) |
| `--cell-address` | Single cell address (e.g., 'B5') - expands to contiguous data region around this cell (required for: get-current-region) (valid for: get-current-region) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### rangeedit

Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data

**Actions:** `insert-cells`, `delete-cells`, `insert-rows`, `delete-rows`, `insert-columns`, `delete-columns`, `find`, `replace`, `sort`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range (required) |
| `--range` | Cell range address where cells will be inserted (e.g., 'A1:D10') (required) |
| `--insert-shift` | Direction to shift existing cells: 'Down' or 'Right' (required for: insert-cells) (valid for: insert-cells) |
| `--delete-shift` | Direction to shift remaining cells: 'Up' or 'Left' (required for: delete-cells) (valid for: delete-cells) |
| `--search-value` | Text or value to search for (required for: find) (valid for: find) |
| `--find-options` | Search options: matchCase (default: false), matchEntireCell (default: false), searchFormulas (default: true) (required for: find) (valid for: find) |
| `--find-value` | Text or value to search for (required for: replace) (valid for: replace) |
| `--replace-value` | Text or value to replace matches with (required for: replace) (valid for: replace) |
| `--replace-options` | Replace options: matchCase (default: false), matchEntireCell (default: false), replaceAll (default: true) (required for: replace) (valid for: replace) |
| `--sort-columns` | Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range (required for: sort) (valid for: sort) (JSON format) |
| `--has-headers` | Whether the range has a header row to exclude from sorting (default: true) (valid for: sort) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### rangeformat

Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions

**Actions:** `set-style`, `get-style`, `format-range`, `format-ranges`, `validate-range`, `get-validation`, `remove-validation`, `auto-fit-columns`, `auto-fit-rows`, `merge-cells`, `unmerge-cells`, `get-merge-info`, `set-column-width`, `set-row-height`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range (required) |
| `--range` | Cell range address (e.g., 'A1:D10') (required for: set-style, get-style, format-range, validate-range, get-validation, remove-validation, auto-fit-columns, auto-fit-rows, merge-cells, unmerge-cells, get-merge-info, set-column-width, set-row-height) (valid for: set-style, get-style, format-range, validate-range, get-validation, remove-validation, auto-fit-columns, auto-fit-rows, merge-cells, unmerge-cells, get-merge-info, set-column-width, set-row-height) |
| `--style-name` | Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent'). Use 'Normal' to reset. (required for: set-style) (valid for: set-style) |
| `--font-name` | Font family name (e.g., 'Arial', 'Calibri', 'Times New Roman') (valid for: format-range, format-ranges) |
| `--font-size` | Font size in points (e.g., 10, 11, 12, 14, 16) (valid for: format-range, format-ranges) |
| `--bold` | Whether to apply bold formatting (valid for: format-range, format-ranges) |
| `--italic` | Whether to apply italic formatting (valid for: format-range, format-ranges) |
| `--underline` | Whether to apply underline formatting (valid for: format-range, format-ranges) |
| `--font-color` | Font (foreground) color as hex '#RRGGBB' (e.g., '#FF0000' for red) (valid for: format-range, format-ranges) |
| `--fill-color` | Cell fill (background) color as hex '#RRGGBB' (e.g., '#FFFF00' for yellow) (valid for: format-range, format-ranges) |
| `--border-style` | Border line style: 'continuous', 'dash', 'dot', 'dashdot', 'dashdotdot', 'double', 'slantdashdot', 'none' (valid for: format-range, format-ranges) |
| `--border-color` | Border color as hex '#RRGGBB' (valid for: format-range, format-ranges) |
| `--border-weight` | Border weight: 'hairline', 'thin', 'medium', 'thick' (valid for: format-range, format-ranges) |
| `--horizontal-alignment` | Horizontal text alignment: 'left', 'center', 'right', 'justify', 'fill' (valid for: format-range, format-ranges) |
| `--vertical-alignment` | Vertical text alignment: 'top', 'center' (or 'middle'), 'bottom', 'justify' (valid for: format-range, format-ranges) |
| `--wrap-text` | Whether to wrap text within cells (valid for: format-range, format-ranges) |
| `--orientation` | Text rotation in degrees (-90 to 90, or 255 for vertical) (valid for: format-range, format-ranges) |
| `--range-addresses` | Cell range addresses to format (e.g., 'A1:D1', 'A3:D3') (required for: format-ranges) (valid for: format-ranges) |
| `--number-format` | Excel number format code applied to all target ranges (e.g., '0.00%' for percentage, '$#,##0.00' for currency, 'm/d/yyyy' for date). LLMs know Excel format codes natively. (valid for: format-ranges) |
| `--validation-type` | Data validation type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom' (required for: validate-range) (valid for: validate-range) |
| `--validation-operator` | Validation comparison operator: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual' (valid for: validate-range) |
| `--formula1` | First validation formula/value - for list validation use range '=$A$1:$A$10' or inline '"A,B,C"' (valid for: validate-range) |
| `--formula2` | Second validation formula/value - required only for 'between' and 'notBetween' operators (valid for: validate-range) |
| `--show-input-message` | Whether to show input message when cell is selected (default: false) (valid for: validate-range) |
| `--input-title` | Title for the input message popup (valid for: validate-range) |
| `--input-message` | Text for the input message popup (valid for: validate-range) |
| `--show-error-alert` | Whether to show error alert on invalid input (default: true) (valid for: validate-range) |
| `--error-style` | Error alert style: 'stop' (prevents entry), 'warning' (allows override), 'information' (allows entry) (valid for: validate-range) |
| `--error-title` | Title for the error alert popup (valid for: validate-range) |
| `--error-message` | Text for the error alert popup (valid for: validate-range) |
| `--ignore-blank` | Whether to allow blank cells in validation (default: true) (valid for: validate-range) |
| `--show-dropdown` | Whether to show dropdown arrow for list validation (default: true) (valid for: validate-range) |
| `--column-width` | Width in points (1 point = 1/72 inch, approx 0.35mm). Standard width ~8.43 points. Range: 0.25-409 points. (required for: set-column-width) (valid for: set-column-width) |
| `--row-height` | Height in points (1 point = 1/72 inch, approx 0.35mm). Default row height ~15 points. Range: 0-409 points. (required for: set-row-height) (valid for: set-row-height) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### rangelink

Hyperlink, threaded comment, and cell protection operations for Excel ranges

**Actions:** `add-hyperlink`, `update-hyperlink`, `remove-hyperlink`, `list-hyperlinks`, `get-hyperlink`, `add-threaded-comment`, `list-threaded-comments`, `add-threaded-comment-reply`, `delete-threaded-comment`, `set-cell-lock`, `get-cell-lock`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet (required) |
| `--cell-address` | Single cell address (e.g., 'A1') (required for: add-hyperlink, update-hyperlink, get-hyperlink, add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment) (valid for: add-hyperlink, update-hyperlink, get-hyperlink, add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment) |
| `--url` | Optional external URL or file path. Omit for an internal workbook link. (valid for: add-hyperlink, update-hyperlink) |
| `--display-text` | Text to display in the cell (optional, defaults to URL) (valid for: add-hyperlink, update-hyperlink) |
| `--tooltip` | Tooltip text shown on hover (optional) (valid for: add-hyperlink, update-hyperlink) |
| `--sub-address` | Optional internal workbook target such as "'Sheet2'!A1" (valid for: add-hyperlink, update-hyperlink) |
| `--range` | Cell range address to remove hyperlinks from (e.g., 'A1:D10') (required for: remove-hyperlink, set-cell-lock, get-cell-lock) (valid for: remove-hyperlink, set-cell-lock, get-cell-lock) |
| `--text` | (required for: add-threaded-comment, add-threaded-comment-reply) (valid for: add-threaded-comment, add-threaded-comment-reply) |
| `--locked` | Lock status: true = locked (protected when sheet protection enabled), false = unlocked (editable) (required for: set-cell-lock) (valid for: set-cell-lock) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### screenshot

Capture Excel worksheet content as images for visual verification

**Actions:** `capture`, `capture-sheet`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet name (null for active sheet) |
| `--range` | Range to capture (e.g., "A1:F20") (valid for: capture) |
| `--quality` | Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### service

Service lifecycle management: start, stop, status

#### service start

Start the ExcelMCP Service if needed and wait up to 30 seconds for readiness

#### service stop

Stop the CLI service and its pipe-owned Excel processes

#### service status

Show stopped, starting, running, or unresponsive daemon state; readiness checks wait up to 10 seconds

### session

Session management. WORKFLOW: open -> use sessionId -> close (--save to persist). Use --show for IRM/auth prompts

#### session create

Create a new Excel file, open it, and create a session. Add --show for visible Excel

| Parameter | Description |
|-----------|-------------|
| `<file>` | Path to the new Excel file to create |
| `--timeout` | Session open/create and operation timeout in whole seconds (default: 120; range: 10-3600) |
| `--show` | Show the Excel window for IRM/auth prompts instead of running hidden |

#### session open

Open an Excel file and create a session. Add --show for visible Excel

| Parameter | Description |
|-----------|-------------|
| `<file>` | Path to the Excel file to open |
| `--timeout` | Session open and operation timeout in whole seconds (default: 120; range: 10-3600) |
| `--show` | Show the Excel window for IRM/auth prompts instead of running hidden |

#### session close

Close a session. Use --save to persist changes

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID to close |
| `--save` | Save changes before closing |

#### session list

List active sessions; transport failures return unresponsive instead of an empty list

#### session test

Test file existence, validity, openability, and IRM/AIP read-only requirements

| Parameter | Description |
|-----------|-------------|
| `<file>` | Full path to test for existence, validity, openability, and IRM/AIP requirements |

### sheet

Worksheet lifecycle management: create, rename, copy, delete, move, list sheets

**Actions:** `list`, `create`, `rename`, `copy`, `delete`, `move`, `copy-to-file`, `move-to-file`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--file-path` | Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook. (valid for: list, create) |
| `--sheet` | Name for the new worksheet (required for: create, delete, move) (valid for: create, delete, move) |
| `--old-name` | Current name of the worksheet (required for: rename) (valid for: rename) |
| `--new-name` | New name for the worksheet (required for: rename) (valid for: rename) |
| `--source-name` | Name of the source worksheet (required for: copy) (valid for: copy) |
| `--target-name` | Name for the copied worksheet (required for: copy) (valid for: copy) |
| `--before-sheet` | Optional: Name of sheet to position before (valid for: move, copy-to-file, move-to-file) |
| `--after-sheet` | Optional: Name of sheet to position after (valid for: move, copy-to-file, move-to-file) |
| `--source-file` | Full path to the source workbook (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--source-sheet` | Name of the sheet to copy (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--target-file` | Full path to the target workbook (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--target-sheet-name` | Optional: New name for the copied sheet (default: keeps original name) (valid for: copy-to-file) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### slicer

Slicer visual filters for PivotTables and Excel Tables

**Actions:** `create-slicer`, `list-slicers`, `set-slicer-selection`, `delete-slicer`, `create-table-slicer`, `list-table-slicers`, `set-table-slicer-selection`, `delete-table-slicer`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable to create slicer for (required for: create-slicer) (valid for: create-slicer, list-slicers) |
| `--field-name` | Name of the field to use for the slicer (required for: create-slicer) (valid for: create-slicer) |
| `--slicer-name` | Name for the new slicer (required for: create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) (valid for: create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) |
| `--destination-sheet` | Worksheet where slicer will be placed (required for: create-slicer, create-table-slicer) (valid for: create-slicer, create-table-slicer) |
| `--position` | Top-left cell position for the slicer (e.g., "H2") (required for: create-slicer, create-table-slicer) (valid for: create-slicer, create-table-slicer) |
| `--selected-items` | Items to select (show in PivotTable) (required for: set-slicer-selection, set-table-slicer-selection) (valid for: set-slicer-selection, set-table-slicer-selection) (JSON format) |
| `--clear-first` | If true, clears existing selection before setting new items (default: true) (valid for: set-slicer-selection, set-table-slicer-selection) |
| `--table-name` | Name of the Excel Table (required for: create-table-slicer) (valid for: create-table-slicer, list-table-slicers) |
| `--column-name` | Name of the column to use for the slicer (required for: create-table-slicer) (valid for: create-table-slicer) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### table

Excel Tables (ListObjects) - lifecycle and data operations

**Actions:** `list`, `create`, `rename`, `delete`, `read`, `resize`, `toggle-totals`, `set-column-total`, `append`, `get-data`, `set-style`, `add-to-data-model`, `create-from-dax`, `update-dax`, `get-dax`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet to create the table on (required for: create, create-from-dax) (valid for: create, create-from-dax) |
| `--table-name` | Name for the new table (must be unique in workbook) (required for: create, rename, delete, read, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-data-model, create-from-dax, update-dax, get-dax) (valid for: create, rename, delete, read, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-data-model, create-from-dax, update-dax, get-dax) |
| `--range` | Cell range address for the table (e.g., 'A1:D10') (required for: create) (valid for: create) |
| `--has-headers` | True if first row contains column headers (default: true) (valid for: create) |
| `--table-style` | Table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1'). Optional. (required for: set-style) (valid for: create, set-style) |
| `--new-name` | New name for the table (must be unique in workbook) (required for: rename) (valid for: rename) |
| `--new-range` | New range address (e.g., 'A1:F20') (required for: resize) (valid for: resize) |
| `--show-totals` | True to show totals row, false to hide (required for: toggle-totals) (valid for: toggle-totals) |
| `--column-name` | Name of the column to set total function on (required for: set-column-total) (valid for: set-column-total) |
| `--total-function` | Totals function name: Sum, Count, Average, Min, Max, CountNums, StdDev, Var, None (required for: set-column-total) (valid for: set-column-total) |
| `--rows` | 2D array of row data to append - column order must match table columns. Optional if rowsFile is provided. (valid for: append) (JSON format) |
| `--rows-file` | Path to a JSON or CSV file containing the rows to append. JSON: 2D array. CSV: rows/columns. Alternative to inline rows parameter. (valid for: append) |
| `--visible-only` | True to return only visible (non-filtered) rows; false for all rows (default: false) (valid for: get-data) |
| `--strip-bracket-column-names` | When true, renames source table columns that contain literal bracket characters (removes brackets) beforeadding to the Data Model. This modifies the Excel table column headers in the worksheet. (valid for: add-to-data-model) |
| `--dax-query` | DAX EVALUATE query (e.g., 'EVALUATE Sales' or 'EVALUATE SUMMARIZE(...) ') (required for: create-from-dax, update-dax) (valid for: create-from-dax, update-dax) |
| `--target-cell` | Target cell address for table placement (default: 'A1') (valid for: create-from-dax) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### tablecolumn

Table column, filtering, and sorting operations for Excel Tables (ListObjects)

**Actions:** `apply-filter`, `apply-filter-values`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `get-column-number-format`, `set-column-number-format`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--table-name` | Name of the Excel table (required) |
| `--column-name` | Name of the column to filter (required for: apply-filter, apply-filter-values, add-column, remove-column, sort, get-column-number-format, set-column-number-format) (valid for: apply-filter, apply-filter-values, add-column, remove-column, get-structured-reference, sort, get-column-number-format, set-column-number-format) |
| `--criteria` | Filter criteria string (e.g., '>100', '=Active', '<>Closed') (required for: apply-filter) (valid for: apply-filter) |
| `--values` | List of exact values to include in the filter (required for: apply-filter-values) (valid for: apply-filter-values) (JSON format) |
| `--position` | 1-based column position (optional, defaults to end of table) (valid for: add-column) |
| `--old-name` | Current column name (required for: rename-column) (valid for: rename-column) |
| `--new-name` | New column name (required for: rename-column) (valid for: rename-column) |
| `--region` | Table region: 'Data', 'Headers', 'Totals', or 'All' (required for: get-structured-reference) (valid for: get-structured-reference) |
| `--ascending` | Sort order: true = ascending (A-Z, 0-9), false = descending (default: true) (valid for: sort) |
| `--sort-columns` | List of sort specifications: [{columnName: 'Col1', ascending: true}, ...] - applied in order (required for: sort-multi) (valid for: sort-multi) (JSON format) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00', '0%', 'yyyy-mm-dd') (required for: set-column-number-format) (valid for: set-column-number-format) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### vba

VBA module and procedure operations for macro-enabled workbooks (.xlsm)

**Actions:** `list`, `view`, `import`, `update`, `run`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--module-name` | Name of the VBA module (required for: view, import, update, delete) (valid for: view, import, update, delete) |
| `--vba-code` | VBA code. Public callers must supply either inline vbaCode or a readable vbaCodeFile, not both. (required for: import, update) (valid for: import, update) |
| `--vba-code-file` | Path to a readable file containing vbaCode; use instead of inline vbaCode, not together (valid for: import, update) |
| `--procedure-name` | Name of the procedure to run (for example "Module1.MySub") (required for: run) (valid for: run) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: run) |
| `--parameters` | Optional parameters to pass to the procedure (valid for: run) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### window

Control Excel window visibility, position, state, status bar, and worksheet-specific views

**Actions:** `show`, `hide`, `bring-to-front`, `get-info`, `set-state`, `set-position`, `arrange`, `set-status-bar`, `clear-status-bar`, `get-view`, `freeze-panes`, `unfreeze-panes`, `set-split`, `set-zoom`, `set-display-options`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--window-state` | Window state: 'normal', 'minimized', or 'maximized' (required for: set-state) (valid for: set-state) |
| `--left` | Window left position in points (valid for: set-position) |
| `--top` | Window top position in points (valid for: set-position) |
| `--width` | Window width in points (valid for: set-position) |
| `--height` | Window height in points (valid for: set-position) |
| `--preset` | Preset name: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen' (required for: arrange) (valid for: arrange) |
| `--text` | Status bar text to display (e.g. "Building PivotTable from Sales data...") (required for: set-status-bar) (valid for: set-status-bar) |
| `--sheet` | Worksheet whose view should be inspected (required for: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, set-display-options) (valid for: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, set-display-options) |
| `--frozen-rows` | Number of rows to freeze from the top (0-1,048,575) (valid for: freeze-panes) |
| `--frozen-columns` | Number of columns to freeze from the left (0-16,383) (valid for: freeze-panes) |
| `--split-rows` | Number of rows above the horizontal split (0-1,048,575) (valid for: set-split) |
| `--split-columns` | Number of columns left of the vertical split (0-16,383) (valid for: set-split) |
| `--zoom` | Zoom percentage from 10 through 400 (required for: set-zoom) (valid for: set-zoom) |
| `--show-gridlines` | Whether to display cell gridlines (valid for: set-display-options) |
| `--show-headings` | Whether to display row and column headings (valid for: set-display-options) |
| `--show-outline-symbols` | Whether to display outline level symbols (valid for: set-display-options) |
| `--show-formulas` | Whether to display formulas instead of their calculated values (valid for: set-display-options) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### workbook

Manage workbook metadata, integrity validation, document properties, Save As/copy operations, fixed-format exports, and external Excel links

**Actions:** `get-info`, `validate-integrity`, `list-document-properties`, `get-document-property`, `set-document-property`, `delete-document-property`, `save-as`, `save-copy-as`, `export-fixed-format`, `list-external-links`, `update-external-link`, `break-external-link`, `set-protection`, `get-protection`, `set-view-options`, `get-view-options`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--checks` | Checks to run; omit for formulas, links, tables, and supplied control totals (valid for: validate-integrity) (JSON format) |
| `--worksheet-names` | Optional worksheet names limiting formula and table checks (valid for: validate-integrity) (JSON format) |
| `--control-totals` | Expected numeric cells with optional absolute tolerances (valid for: validate-integrity) (JSON format) |
| `--max-findings` | Maximum finding details to return; counts still include omitted details. Range: 1-10000. (valid for: validate-integrity) |
| `--include-built-in` | (valid for: list-document-properties) |
| `--include-custom` | (valid for: list-document-properties) |
| `--property-name` | Document property name (required for: get-document-property, set-document-property, delete-document-property) (valid for: get-document-property, set-document-property, delete-document-property) |
| `--scope` | Property collection: built-in or custom (valid for: get-document-property, set-document-property) |
| `--value` | String value to store (required for: set-document-property) (valid for: set-document-property) |
| `--target-path` | Absolute output path in an existing directory (required for: save-as, save-copy-as, export-fixed-format) (valid for: save-as, save-copy-as, export-fixed-format) |
| `--format` | Output format: auto, xlsx, xlsm, xlsb, or xls (valid for: save-as) |
| `--overwrite` | Whether an existing output file may be replaced (valid for: save-as, save-copy-as, export-fixed-format) |
| `--format-type` | (valid for: export-fixed-format) |
| `--quality` | (valid for: export-fixed-format) |
| `--include-document-properties` | (valid for: export-fixed-format) |
| `--ignore-print-areas` | (valid for: export-fixed-format) |
| `--from-page` | (valid for: export-fixed-format) |
| `--to-page` | (valid for: export-fixed-format) |
| `--open-after-publish` | (valid for: export-fixed-format) |
| `--link-source` | (required for: update-external-link, break-external-link) (valid for: update-external-link, break-external-link) |
| `--is-protected` | (required for: set-protection) (valid for: set-protection) |
| `--password` | (valid for: set-protection) |
| `--display-gridlines` | (valid for: set-view-options) |
| `--display-headings` | (valid for: set-view-options) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### worksheetstyle

Worksheet styling, visibility, protection, grouping, and outline operations

**Actions:** `set-tab-color`, `get-tab-color`, `clear-tab-color`, `set-protection`, `get-protection`, `set-comment`, `get-comment`, `clear-comment`, `add-image`, `get-image-count`, `add-shape`, `get-shape-count`, `set-page-setup`, `get-page-setup`, `set-visibility`, `get-visibility`, `show`, `hide`, `very-hide`, `group`, `ungroup`, `get-outline-info`, `set-outline-settings`, `show-outline-levels`, `clear-outline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet to color (required) |
| `--red` | Red color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--green` | Green color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--blue` | Blue color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--is-protected` | Whether the worksheet should be protected (required for: set-protection) (valid for: set-protection) |
| `--password` | Optional password for protecting/unprotecting the sheet (valid for: set-protection) |
| `--cell-address` | Cell address such as A1 (required for: set-comment, get-comment, clear-comment, add-image, add-shape) (valid for: set-comment, get-comment, clear-comment, add-image, add-shape) |
| `--text` | Cell note text to set (required for: set-comment) (valid for: set-comment) |
| `--image-path` | Absolute path to the image file on disk (required for: add-image) (valid for: add-image) |
| `--orientation` | Page orientation: 'portrait' or 'landscape' (required for: set-page-setup) (valid for: set-page-setup) |
| `--fit-to-pages-wide` | Number of pages wide to fit the printout to (valid for: set-page-setup) |
| `--fit-to-pages-tall` | Number of pages tall to fit the printout to (valid for: set-page-setup) |
| `--center-horizontally` | Whether to center the printout horizontally on the page (valid for: set-page-setup) |
| `--center-vertically` | Whether to center the printout vertically on the page (valid for: set-page-setup) |
| `--visibility` | Visibility level: 'visible', 'hidden', or 'veryhidden' (required for: set-visibility) (valid for: set-visibility) |
| `--range` | Row or column range to group (required for: group, ungroup, get-outline-info) (valid for: group, ungroup, get-outline-info) |
| `--axis` | Grouping axis: Rows or Columns (required for: group, ungroup, get-outline-info) (valid for: group, ungroup, get-outline-info) |
| `--summary-row` | Summary row position: above or below (valid for: set-outline-settings) |
| `--summary-column` | Summary column position: left or right (valid for: set-outline-settings) |
| `--automatic-styles` | Whether Excel applies automatic outline styles (valid for: set-outline-settings) |
| `--row-levels` | Optional row outline level to display (valid for: show-outline-levels) |
| `--column-levels` | Optional column outline level to display (valid for: show-outline-levels) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

### xmlmap

Manage workbook XML maps and exchange XML data without interactive dialogs

**Actions:** `list`, `add`, `map-range`, `import-xml`, `export-xml`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--schema` | XSD schema content. Public callers must supply either inline schema or a readable schemaFile, not both. (required for: add) (valid for: add) |
| `--schema-file` | Path to a readable file containing schema; use instead of inline schema, not together (valid for: add) |
| `--root-element-name` | Optional root element when the schema has multiple roots (valid for: add) |
| `--map-name` | Optional name to assign to the created map (required for: map-range, export-xml, delete) (valid for: add, map-range, import-xml, export-xml, delete) |
| `--sheet` | Worksheet containing the target range (required for: map-range) (valid for: map-range, import-xml) |
| `--range` | Target cell or single-column range (required for: map-range) (valid for: map-range) |
| `--xpath` | XPath to map (required for: map-range) (valid for: map-range) |
| `--selection-namespace` | Optional namespace declarations used by prefixed XPath expressions (valid for: map-range) |
| `--repeating` | Whether to create a repeating XML list mapping (valid for: map-range) |
| `--xml-data` | XML data. Public callers must supply either inline xmlData or a readable xmlDataFile, not both. (required for: import-xml) (valid for: import-xml) |
| `--xml-data-file` | Path to a readable file containing xmlData; use instead of inline xmlData, not together (valid for: import-xml) |
| `--start-cell` | Top-left destination cell for automatic mapping (valid for: import-xml) |
| `--overwrite` | Whether imported XML may overwrite mapped cells (valid for: import-xml) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

## Common Pitfalls

- `--values-file` requires an existing JSON or CSV file; use `--values` for inline JSON.
- `--timeout` ranges are action-specific: session open/create accepts 10-3600; Power Query refresh/refresh-all accepts 0-2147483 (0 keeps the default); other generated timeout actions accept 1-2147483.
- `pythoninexcel get-result --max-wait-seconds` must be at least 1 and shorter than the session operation timeout.
- `--values` and list parameters use JSON arrays; range values use a two-dimensional array.
- Power Query operations may take 30 seconds or longer; use a deliberate data-operation timeout or 0 for the default.

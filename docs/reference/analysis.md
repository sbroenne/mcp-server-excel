# analysis - What-If Analysis

Use `analysis` for Excel's native Goal Seek, scenarios, scenario summaries, and one- or two-variable data tables.

## Goal Seek

The formula cell must contain a formula, and the changing cell must be one of its inputs.

```mcp
analysis(action: 'goal-seek', session_id: sessionId, sheet_name: 'Model', formula_cell: 'B10', goal: 10000, changing_cell: 'B3')
```

```cli
excelcli -q analysis goal-seek --session $sessionId --sheet Model --formula-cell B10 --goal 10000 --changing-cell B3
```

Goal Seek changes the workbook immediately. Read both cells afterward when the exact final values matter.

## Scenarios

Scenario values must contain exactly one value per cell in `changing_cells`
(MCP) / `--changing-cells` (CLI), in range order.
Showing a scenario replaces those inputs; listing scenarios does not authorize
showing one during an audit.

```mcp
analysis(action: 'create-scenario', session_id: sessionId, sheet_name: 'Model', scenario_name: 'Growth', changing_cells: 'B3:B5', values: ['0.08', '1200', '0.35'])
analysis(action: 'show-scenario', session_id: sessionId, sheet_name: 'Model', scenario_name: 'Growth')
```

```cli
excelcli -q analysis create-scenario --session $sessionId --sheet Model --scenario-name Growth --changing-cells B3:B5 --values '[0.08,1200,0.35]'
excelcli -q analysis show-scenario --session $sessionId --sheet Model --scenario-name Growth
```

Use `create-scenario-summary` when a summary is requested after defining the
scenarios. Use `report_type` (MCP) / `--report-type` (CLI): `summary` for a
normal report sheet or `pivot-table` for a Scenario PivotTable.
Use `result_cells` (MCP) / `--result-cells` (CLI) to identify formulas that depend on the changing cells.

## Data Tables

Prepare the worksheet layout first, including the formula in the table's corner and the input values along its first row or column.

- One-variable row table: provide `row_input_cell` (MCP) / `--row-input-cell` (CLI).
- One-variable column table: provide `column_input_cell` (MCP) / `--column-input-cell` (CLI).
- Two-variable table: provide both.

```mcp
analysis(action: 'create-data-table', session_id: sessionId, sheet_name: 'Model', table_range: 'A1:B11', column_input_cell: 'D1')
```

```cli
excelcli -q analysis create-data-table --session $sessionId --sheet Model --table-range A1:B11 --column-input-cell D1
```

Data tables can be calculation-intensive. Use `calculation_mode` (MCP) /
`calculationmode` (CLI) and follow the shared
[calculation-mode rules](behavioral-rules.md#changes-and-formatting) when
controlling recalculation around larger workbook edits.

## Solver Is Not Exposed

Solver is an optional VBA add-in, not an Excel PIA API. Microsoft requires users to enable the add-in in Excel Options and establish a VBA reference before calling Solver functions. Do not try to invoke Solver through `vba`, enable the add-in, or change macro-security settings automatically. Use Goal Seek for one-variable targets or document that multi-variable constrained optimization requires user-configured Solver.

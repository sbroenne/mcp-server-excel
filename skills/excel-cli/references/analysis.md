> **CLI syntax note:** This shared domain guide may use MCP-style `tool(action: ...)` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live `--help`; notably, MCP `file` open/close maps to CLI `session` open/close, and MCP `worksheet` maps to CLI `sheet`.

# analysis - What-If Analysis

Use `analysis` for Excel's native Goal Seek, scenarios, scenario summaries, and one- or two-variable data tables.

## Goal Seek

The formula cell must contain a formula, and the changing cell must be one of its inputs.

```text
analysis(action="goal-seek", sheet_name="Model", formula_cell="B10", goal=10000, changing_cell="B3")
```

Goal Seek changes the workbook immediately. Read both cells afterward when the exact final values matter.

## Scenarios

Scenario values must contain exactly one value per cell in `changing_cells`, in range order.

```text
analysis(action="create-scenario", sheet_name="Model", scenario_name="Growth",
         changing_cells="B3:B5", values=[0.08, 1200, 0.35])
analysis(action="show-scenario", sheet_name="Model", scenario_name="Growth")
analysis(action="list-scenarios", sheet_name="Model")
```

Use `create-scenario-summary` after defining two or more scenarios. Set `report_type` to `summary` for a normal report sheet or `pivot-table` for a Scenario PivotTable. `result_cells` should identify formulas that depend on the changing cells.

## Data Tables

Prepare the worksheet layout first, including the formula in the table's corner and the input values along its first row or column.

- One-variable row table: provide `row_input_cell`.
- One-variable column table: provide `column_input_cell`.
- Two-variable table: provide both.

```text
analysis(action="create-data-table", sheet_name="Model", table_range="A1:B11", column_input_cell="D1")
```

Data tables can be calculation-intensive. Use `calculation_mode` when controlling recalculation around larger workbook edits.

## Solver Is Not Exposed

Solver is an optional VBA add-in, not an Excel PIA API. Microsoft requires users to enable the add-in in Excel Options and establish a VBA reference before calling Solver functions. Do not try to invoke Solver through `vba`, enable the add-in, or change macro-security settings automatically. Use Goal Seek for one-variable targets or document that multi-variable constrained optimization requires user-configured Solver.

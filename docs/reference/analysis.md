# What-If Analysis

Goal Seek, scenarios, and data tables answer different questions. Choose the
method from the requested result, and use CLI help or MCP tool descriptions
for current commands and inputs.

**Mac experimental beta:** Goal Seek and one-/two-variable Data Tables are
enabled; all Scenario actions below remain Windows-only. See
[macOS support](https://excelmcpserver.dev/macos-support/).

## Goal Seek

Use Goal Seek when one formula result must reach a numeric target by changing
one input. The changing cell must actually influence the formula.

Goal Seek changes live cells. It is not a read-only diagnostic; inspect both
the resulting input and formula value before treating the target as established.

## Scenarios

Use scenarios to compare named sets of assumptions. Match each value to its
intended input cell in the correct order.

Applying a scenario replaces those inputs. Listing scenarios during an audit
does not authorize applying one. After the chosen assumptions are configured,
request a summary when the task needs a comparison report.

## Data Tables

Prepare the layout before creating a sensitivity table: the formula, trial
values, and input cells must represent the intended one- or two-variable
comparison. A What-If data table is not an ordinary worksheet Table.

Large tables can be expensive to calculate. If controlling recalculation during
edits, follow [calculation guidance](calculation.md) and restore the original
mode after failure as well as success.

## Solver Is Not Exposed

Solver is an optional VBA add-in that needs separate user configuration and
macro-security decisions. Do not enable it or change trust settings automatically.
Use Goal Seek for a one-variable target, or explain when the request needs
user-configured constrained optimization.

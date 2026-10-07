# Power Query: loading and recovery

Use CLI help or MCP tool descriptions for current commands and inputs. This
guide explains execution, refresh order, and recovery decisions.

## Execution and destination decisions

For authorized development, prefer `evaluate` for new or materially changed M:
check the preview before storing it. Skip redundant evaluation for trivial
literal tables or validated code with unchanged dependencies.

Evaluation is not a read-only operation: it creates temporary workbook objects,
executes M, then removes them by exact identity. Cleanup failure is an error.
Execution may contact external sources. For an audit, inspect definitions and
existing loaded values; follow [permission rules](behavioral-rules.md#intent-and-permission).

| Decision | Effect |
|----------|--------|
| Create a query | Stores it and can immediately execute its chosen load |
| Replace an existing query's code | Can refresh its data; choose definition-only editing when execution is not intended |
| Change a load destination | Loads immediately, not just configuration |
| Refresh an existing load | Updates data; definitions alone do not prove loaded values are current |
| Remove loads but keep the query | Removes all worksheet/model destinations |
| Remove the query | Also removes associated loads; inspect dependencies first |

Use `connection-only` to store without execution. It does not validate M.
A query loaded only to the model is not connection-only.

When creating/loading onto populated sheets, inspect cells and choose an
explicit destination that preserves other content. Existing Tables
refresh in place. Moving a load requires unload/reload, which removes **all**
current destinations; check model dependencies and authorized scope first.

Set explicit M column types for dates and relationship keys. Query refresh
synchronizes its model load; refresh dependent PivotTables afterward.
See [Data Model](datamodel.md).

## Recovering a failed create

Create adds the query before loading; a load failure can leave the query and
load objects behind. Repeating create then fails with "already exists".
Inspect `list`, `view`, and `get-load-config`; update the surviving query.
Create only if it is absent. Do not blindly delete/rebuild it.

For surviving `SalesQuery` and corrected code in a known `query.m` file:

```mcp
powerquery(action: 'evaluate', workbook_session_id: sessionId, m_code_file: 'query.m')
powerquery(action: 'update', workbook_session_id: sessionId, query_name: 'SalesQuery', m_code_file: 'query.m', refresh: false)
powerquery_read(action: 'get-load-config', workbook_session_id: sessionId, query_name: 'SalesQuery')
```

```cli
excelcli -q powerquery evaluate --session $sessionId --m-code-file query.m
excelcli -q powerquery update --session $sessionId --query-name SalesQuery --m-code-file query.m --refresh false
excelcli -q powerquery get-load-config --session $sessionId --query-name SalesQuery
```

Check each result. Refresh the surviving intended load, or use `load-to` after
checking its destination. Successful evaluation does not prove a sheet load
will succeed. Read affected loaded values before saving.

Cleanup uses the exact case-insensitive mashup `Location`, not display-name
prefixes. Queries such as `A` and `AA` are independent; do not remove similarly
named connections as a shortcut.

## Code, reads, and waits

### Staging queries and batch refresh

`refresh-all` refreshes every loaded query in the workbook. Parameter queries
and definition-only staging queries have no connection or table of their own,
so they are skipped and listed in `skippedQueries`; Excel evaluates them when
the loaded queries that use them refresh. Do not load a staging query onto a
sheet or into the model merely to make batch refresh pass.

When one query fails, `refresh-all` records it in `failedQueries` (with its
error message and category) and keeps refreshing the rest. The result then has
`success: false`, and the CLI exits with code 1. This is not a transaction:
queries listed in `refreshedQueries` keep their new data. Engine errors are
reported, not hidden or converted into success. A timeout, cancellation, or
lost Excel connection stops the whole run instead.

To refresh only selected queries, inspect `get-load-config` to identify loaded
destinations and refresh each intended loaded query by name. Refresh dependent
PivotTables separately afterward.

For a staging definition used by the already-loaded `Financials` query:

```mcp
powerquery_read(action: 'get-load-config', workbook_session_id: sessionId, query_name: 'Financials')
powerquery(action: 'refresh', workbook_session_id: sessionId, query_name: 'Financials')
pivottable(action: 'refresh', workbook_session_id: sessionId, pivot_table_name: 'RevenuePivot')
```

```cli
excelcli -q powerquery get-load-config --session $sessionId --query-name Financials
excelcli -q powerquery refresh --session $sessionId --query-name Financials
excelcli -q pivottable refresh --session $sessionId --pivot-table-name RevenuePivot
```

Check each result and actual loaded values. Repeat for each requested loaded
query and dependent pivot. Refresh executes the stored M sources; it does not
discover new reports or replace a frozen data snapshot with current data.

### Reads and execution limits {#inputs-and-timeouts}

A compact query listing is not the full stored definition, and neither a
definition nor its load settings proves freshness. Inspect the full code when
its sources or transformations matter. Renaming does not rewrite dependent
M references; see [M identifiers and query chaining](m-code-syntax.md).

Remote formatting sends M code to an external service and needs consent.
It is not a workaround for query-engine errors.

Session and refresh waits serve different purposes. Use current CLI/MCP help
for the operation's limits, then inspect surviving objects after timeout.
Extending a wait is not a fix for a query with no refreshable destination.

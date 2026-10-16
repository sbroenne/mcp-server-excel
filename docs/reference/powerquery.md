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

## Power Query tracing

Tracing is controlled by a setting for the current Windows user and does not
belong to a workbook or session. Enable it with:

```cli
excelcli settings set-pq-tracing --enabled true
```

```mcp
settings(action: 'set-pq-tracing', enabled: true)
```

Restart Excel for a tracing change to take effect. The setting stays as chosen
until you change it; closing a workbook, session, or ExcelMCP service does not
restore it. Any Excel process started while tracing is on may write trace files.

Troubleshoot one workbook session at a time. There is no need to empty the
trace folder: filter by the workbook session and refresh time instead. One
Excel process can produce several files, including Power Query helper-process
logs; the session filter matches the Excel process ID and start time, so a
recycled process ID does not pull in earlier files.

Use `trace: true` on Power Query `refresh` or `refresh-all` to attach a compact
summary for that refresh. For existing traces, use `powerquery_trace_read` to list or summarize files:

```mcp
powerquery(action: 'refresh', query_name: 'SalesData', trace: true)
powerquery_trace_read(action: 'summarize', workbook_session_id: sessionId)
```

```cli
excelcli powerquery refresh --session $session --query-name SalesData --trace
excelcli powerquerytrace summarize --workbook-session-id $session --since 15m
```

If optional trace collection fails, the refresh outcome is preserved.
`traceSummary.success` is false and `traceSummary.errorMessage` explains the
trace-read failure; `traceNote` distinguishes that failure from the refresh.
Do not repeat a completed refresh merely to collect traces.

Summaries hide native query text and exception messages by default. Set
`include_query_text: true` (CLI: `--include-query-text`) only when those details
are needed; they can contain sensitive information. Trace files are stored by
Excel under the current user's local application data folder. They may contain
source names, query text, and error details. Manage or remove these files
outside ExcelMcp. Use the returned `traceCursor` as `since` on the next summary
to inspect only newer entries.

The global setting can also be read with `settings_read` action
`get-pq-tracing` and changed with `settings` action `set-pq-tracing`; CLI
commands are `excelcli settings get-pq-tracing` and
`excelcli settings set-pq-tracing --enabled true`.

## Workbook privacy setting

Use `get-privacy-settings` to read this workbook's Power Query privacy behavior
and `set-privacy-settings` to change it. Set `ignorePrivacyLevels` (CLI flag:
`--ignore-privacy-levels`; MCP input: `ignore_privacy_levels`; batch JSON key:
`ignorePrivacyLevels`) to `true` only when authorized. Ignoring privacy levels
can let data from different sources be combined and sent between them. The
response reports the configured workbook setting, not a guarantee of effective
behavior; Excel's global privacy option may override it.

```mcp
powerquery_read(action: 'get-privacy-settings', workbook_session_id: sessionId)
```

```mcp
powerquery(action: 'set-privacy-settings', workbook_session_id: sessionId, ignore_privacy_levels: true)
```

```cli
excelcli powerquery set-privacy-settings --session $sessionId --ignore-privacy-levels true
```

```json
{"command":"powerquery.set-privacy-settings","args":{"ignorePrivacyLevels":true}}
```

This is a workbook-level setting. It is stored when the workbook is saved, and
Excel honors it for users whose global privacy choice is the default
“Combine data according to each file's Privacy Level settings.”

Formula.Firewall failures are reported with
`errorCategory: "PrivacyFirewall"` and point to
`powerquery set-privacy-settings --ignore-privacy-levels true` (MCP action
`set-privacy-settings` with `ignore_privacy_levels: true`). Other privacy
failures keep the `Privacy` category.

## Recovering a failed create

Create adds the query before loading; a load failure can leave the query and
load objects behind. Repeating create then fails with "already exists".
Create and update report recognized Formula.Firewall failures with
`errorCategory: "PrivacyFirewall"` and
`exceptionType: "PowerQueryCommandException"`, like evaluate and refresh.
Other privacy-related errors use `errorCategory: "Privacy"`. The error message
describes the retained changes.
An update can save the replacement M code before its refresh fails: the
stored definition then differs from the data still displayed.

Both MCP and CLI errors include `partialState` after a confirmed definition
change; CLI batch output carries it and the error diagnostics on the failed
item. It contains `queryName`, `queryCreated` for create or `mCodeUpdated`
for update, and the inspected `queryExists`, `loadMode`, `targetSheet` when
present, and `isLoadedToDataModel`. These are the actual destination objects,
not the requested configuration or proof of successful data loading. For
example, a failed `both` create can report `loadMode: "load-to-table"` with
`isLoadedToDataModel: false`.

If inspection fails, `stateInspectionError` explains that separate failure;
unknown state fields are omitted rather than reported as false or
connection-only. Failures before storing a definition have no `partialState`.
Timeout, cancellation, and lost-Excel errors retain their existing recovery
behavior and do not attempt unsafe post-failure inspection.

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

`refresh-all` refreshes every loaded query in the workbook, then refreshes the
PivotTables that depend on those query loads. Shared PivotTable caches are
refreshed once. Set `refresh_pivot_tables=false` (CLI:
`--refresh-pivot-tables false`) to refresh query loads only. Parameter queries
and definition-only staging queries have no connection or table of their own,
so they are skipped and listed in `skippedQueries`; Excel evaluates them when
the loaded queries that use them refresh. Do not load a staging query onto a
sheet or into the model merely to make batch refresh pass.

When one query fails, `refresh-all` records it in `failedQueries` (with its
error message and category) and keeps refreshing the rest. The result then has
`success: false`, and the CLI exits with code 1. When every failed query has the
same category, the result's top-level `errorCategory` repeats it. This is not a
transaction:
queries listed in `refreshedQueries` keep their new data. Engine errors are
reported, not hidden or converted into success. A timeout, cancellation (by the
caller or by Excel), or lost Excel connection stops the whole run instead.

The separate `workbook refresh-all` action matches Excel's workbook-wide
refresh behavior: it refreshes Power Query loads, other workbook data
connections, and all PivotTables. It reports outcomes for queries, connections,
and PivotTables separately. Use it when other workbook connections should be
included; `powerquery refresh-all` does not refresh unrelated connections.

To refresh only selected queries, inspect `get-load-config` to identify loaded
destinations and refresh each intended loaded query by name. Refresh dependent
PivotTables separately afterward when using single-query refresh.

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

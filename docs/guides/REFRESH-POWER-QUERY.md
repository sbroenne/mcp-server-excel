# Refresh Power Query from an AI Assistant

Power Query refreshes are the single most common reason people want Excel
automation: the query logic already works, it just needs to run — on a schedule,
after a source file lands, or as part of a larger task an AI assistant is doing.

ExcelMcp refreshes Power Query by driving the **real Excel application** through
its COM API, so the Power Query engine that runs is Excel's own. Credentials,
privacy levels, native connectors, and the Data Model all behave exactly as they
do when you click **Data → Refresh All** yourself.

## What you ask for

Talk to your assistant in plain language:

> Refresh the `SalesData` query in `Q3-report.xlsx` and tell me how many rows it
> loaded.

> Open `budget.xlsx`, refresh every query, then save it.

Behind the scenes the assistant calls the `powerquery` tool. You do not need to
know the operation names — but they are useful when scripting.

## Refresh one query

=== "MCP Server"

    The assistant opens a session, refreshes, then closes it:

    ```text
    file(action: 'open', file_path: 'C:\reports\Q3-report.xlsx')
    powerquery(action: 'refresh', query_name: 'SalesData', timeout_seconds: 300)
    file(action: 'close', save: true)
    ```

=== "CLI"

    ```powershell
    $session = (excelcli -q session open C:\reports\Q3-report.xlsx | ConvertFrom-Json).data.sessionId
    excelcli -q powerquery refresh --session $session --query-name SalesData --timeout-seconds 300
    excelcli -q session close --session $session --save
    ```

Public timeout values are always integer seconds. `refresh` defaults to a
**30-minute** data-operation timeout when its timeout is omitted or `0`. For
quick queries pass something smaller (60–120 seconds) so a stuck refresh fails
fast instead of holding the session.

## Understand timeout ownership

| Timeout | Configure it with | Applies to | Default and range |
|---|---|---|---|
| Session operation timeout | MCP `file open/create` `timeout_seconds`; CLI `session open/create --timeout-seconds` | Workbook startup and operations that do not supply a dedicated data-operation timeout, including Power Query `create`, `update`, and `evaluate` | 120 seconds; 10–3600 |
| Refresh/data-operation timeout | The refresh action's MCP `timeout_seconds`, CLI `--timeout-seconds`, or batch `timeoutSeconds` | Power Query `refresh` and `refresh-all`; `load-to` uses the same 30-minute default but has no caller override | Power Query omitted/`0`: 1800 seconds; explicit range 0–2147483. Other timeout-bearing categories use 1–2147483. |

A dedicated refresh/data-operation timeout replaces the session operation wait
for that operation; the two limits are not layered. `powerquery load-to` has no
caller timeout option and uses the fixed 30-minute data-operation timeout;
passing `timeoutSeconds` is rejected as action-inapplicable rather than ignored.

## Refresh everything

Use the `refresh-all` action to refresh every query in the workbook:

```powershell
excelcli -q powerquery refresh-all --session $session
```

By default, Power Query `refresh-all` also refreshes PivotTables that depend on
successfully refreshed query loads; shared caches are refreshed once. The result
lists refreshed and failed PivotTables alongside `refreshedQueries`,
`skippedQueries`, and `failedQueries`. Set `refresh_pivot_tables=false` in MCP or
`--refresh-pivot-tables false` in the CLI to refresh queries only.

Parameter and connection-only queries are skipped because they have nothing of
their own to refresh; Excel evaluates them when the loaded queries that use
them refresh. If a query or PivotTable refresh fails, the operation continues
where possible, reports `success: false`, and the CLI exits with code 1.
Successful refreshes keep their new data. For Excel-style workbook-wide refresh
of Power Query and other data connections plus all PivotTables, use
`workbook refresh-all`.

## Include a trace summary

Add `trace: true` to MCP `powerquery` action `refresh` or `refresh-all`, or pass
`--trace` in the CLI. The result includes a compact summary of Power Query's
activity during that refresh, including data-source time and quiet gaps. It does
not restart Excel. If no trace entries are found, enable tracing with MCP
`settings` action `set-pq-tracing` and `enabled=true`, or CLI
`excelcli settings set-pq-tracing --enabled true`, then restart Excel before
refreshing again. This is a per-user setting, so other Excel processes started
while it is on may also write trace files. The setting remains enabled until
you explicitly turn it off.

Keep troubleshooting scoped to one workbook session. A session can write more
than one trace file because Power Query uses helper processes. There is no need
to clear the folder: `trace: true` selects the current Excel process and refresh
time automatically. If reading the trace files fails, the refresh outcome is
unchanged; inspect `traceSummary.errorMessage` and `traceNote` separately.

For traces collected outside a refresh call, use MCP `powerquery_trace_read`
actions `list` and `summarize` or CLI `powerquerytrace list` and
`powerquerytrace summarize`. `traceCursor` can be passed back as `since` to
read only newer entries. Query text and exception messages are omitted unless
`include_query_text: true` is explicitly requested.

## Test M code before you save it

The most valuable habit when an AI assistant writes Power Query for you is
**test-then-commit**. The `evaluate` action runs M code and returns a data
preview *without* creating a permanent query:

```text
1. powerquery(action: 'evaluate', mCode: '...')   # test - nothing is persisted
2. powerquery(action: 'create',   queryName: 'SalesData', mCode: '...')
3. powerquery(action: 'refresh',  queryName: 'SalesData')
```

`evaluate` returns the Power Query engine's real error message, which is far more
useful than the COM exception you get from a failed `create`. Skipping it is how
workbooks end up polluted with broken queries.

For large M programs, use `m_code_file` in MCP, `--m-code-file` in the CLI, or
`mCodeFile` in batch JSON. Supply exactly one of the inline and file forms. The
file must exist and be readable; file content is resolved once by shared
dispatch.

Use `create` for a new query and `update` for one that already exists — `create`
fails with "Query 'X' already exists", and `update` fails with "not found". Run
`powerquery list` first if you are unsure.

## Verify the refresh worked

Always confirm rather than assuming:

```powershell
excelcli -q powerquery list --session $session          # load state per query
excelcli -q powerquery get-load-config --session $session --query-name SalesData
excelcli -q datamodel list-tables --session $session    # if loading to the Data Model
```

In `list` output, `IsConnectionOnly = true` means the query has **no** data
destination. A query loaded only to the Data Model is *not* connection-only.
List includes exact load mode and formula character count, without formula
previews. Use `powerquery view` (MCP: `powerquery_read` action `view`) for one
query's complete M code.
List, View, and Get Load Config share one exact worksheet/Data Model-aware
detector. If Excel cannot inspect a query, List fails instead of silently
omitting it from a successful response.

## Choose where the data lands

The load destination controls where refreshed data goes:

| Value | Result |
|---|---|
| `worksheet` | Creates an Excel Table on a worksheet (default) |
| `data-model` | Loads to Power Pivot for DAX analysis |
| `both` | Loads to the worksheet **and** Power Pivot |
| `connection-only` | Imports the query definition without loading data |

To load onto a sheet that already contains data, pass a target cell address (for
example `B5`) so ExcelMcp places the table instead of clearing the sheet. Omit it
on a populated sheet and the tool returns guidance asking for one.

## Known gotchas

**M code validation only happens on execution.** A `connection-only` query is not
validated until something runs it. Use `evaluate` to check code up front.

**Column names with hyphens, spaces, or punctuation must be quoted.**
`[Non-Recurring]` parses as `[Non] - [Recurring]` — subtraction — and fails with a
confusing "The name 'X' wasn't recognized". Write `[#"Non-Recurring"]`. The rule:
anything other than letters, digits, and underscores needs `[#"..."]`.

**Always set column types explicitly.** Include `Table.TransformColumnTypes()` in
your M code. Without it, dates can be stored as numbers and Data Model
relationships silently fail to match.

**`unload` and `delete` also remove Data Model connections**, not just the
worksheet table. `unload` keeps the query definition; `delete` removes everything.

**Refreshing Power Query also refreshes the Data Model.** Tables loaded via Power
Query auto-sync, so PivotTables connected to the Data Model update too. Worksheet
tables edited directly do *not* — see
[Query the Data Model with DAX](QUERY-DATA-MODEL-WITH-DAX.md).

**Excel dialogs.** Confirmed blocking dialogs automatically reveal the session's
Excel window and report `excelState: "dialogOpen"` in MCP `file_read` action
`list` or CLI `excelcli session list`. Ask the user to check the window and
respond there. Readable titles and message labels are included automatically in
`blockingMessage` and timeout errors; unavailable or partial reads are marked.
The refresh remains pending. The server does not identify dialog types, read
input values, answer dialogs, or change security settings.

**M code is never sent anywhere by default.** `create` and `update` preserve your M
code exactly. Only opting in to remote formatting sends it to the external
powerqueryformatter.com service, and only with explicit consent.

## Related

- [Power Query, DAX and analytics operations](../features/DATA-ANALYTICS.md)
- [Query the Data Model with DAX](QUERY-DATA-MODEL-WITH-DAX.md)
- [Build and update PivotTables with an AI assistant](AUTOMATE-PIVOTTABLES.md)

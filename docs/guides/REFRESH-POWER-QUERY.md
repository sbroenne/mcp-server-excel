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
    file(action: 'open', path: 'C:\reports\Q3-report.xlsx')
    powerquery(action: 'refresh', query_name: 'SalesData', timeout_seconds: 300)
    file(action: 'close', save: true)
    ```

=== "CLI"

    ```powershell
    $session = (excelcli -q session open C:\reports\Q3-report.xlsx | ConvertFrom-Json).sessionId
    excelcli -q powerquery refresh --session $session --query-name SalesData --timeout 300
    excelcli -q session close --session $session --save
    ```

Public timeout values are always integer seconds. `refresh` defaults to a
**30-minute** data-operation timeout when its timeout is omitted or `0`. For
quick queries pass something smaller (60–120 seconds) so a stuck refresh fails
fast instead of holding the session.

## Understand timeout ownership

| Timeout | Configure it with | Applies to | Default and range |
|---|---|---|---|
| Session operation timeout | MCP `file open/create` `timeout_seconds`; CLI `session open/create --timeout` | Workbook startup and operations that do not supply a dedicated data-operation timeout, including Power Query `create`, `update`, and `evaluate` | 120 seconds; 10–3600 |
| Refresh/data-operation timeout | The refresh action's MCP `timeout_seconds`, CLI `--timeout`, or batch `timeout` | Power Query `refresh` and `refresh-all`; `load-to` uses the same 30-minute default but has no caller override | Power Query omitted/`0`: 1800 seconds; explicit range 0–2147483. Other timeout-bearing categories use 1–2147483. |

A dedicated refresh/data-operation timeout replaces the session operation wait
for that operation; the two limits are not layered. `powerquery load-to` has no
caller timeout option and uses the fixed 30-minute data-operation timeout;
passing `timeout` is rejected as action-inapplicable rather than ignored.

## Refresh everything

Use the `refresh-all` action to refresh every query in the workbook:

```powershell
excelcli -q powerquery refresh-all --session $session
```

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
List includes exact load mode, formula character count, and an M preview of at
most 80 characters; use `powerquery view` for one query's complete M code.
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

**Credential and privacy dialogs.** If Excel blocks on a sign-in or privacy prompt,
the operation returns suggested next actions rather than hanging. Surface those to
the user and retry once the prompt is cleared.

**M code is never sent anywhere by default.** `create` and `update` preserve your M
code exactly. Only opting in to remote formatting sends it to the external
powerqueryformatter.com service, and only with explicit consent.

## Related

- [Power Query, DAX and analytics operations](../features/DATA-ANALYTICS.md)
- [Query the Data Model with DAX](QUERY-DATA-MODEL-WITH-DAX.md)
- [Build and update PivotTables with an AI assistant](AUTOMATE-PIVOTTABLES.md)

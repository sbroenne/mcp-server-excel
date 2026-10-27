# Choosing an Excel workflow

These are alternatives, not steps to run on every workbook. Read only the
guide needed for an unresolved decision.

| Need | Surface |
|------|---------|
| Existing cell values or formulas | [Ranges](range.md); no query/model setup |
| Direct CSV or legacy HTML import | [QueryTables](querytable.md) |
| Modern imports or stored transformations | [Power Query](powerquery.md) |
| Aggregation over loaded model data | [Data Model/DAX](datamodel.md) |
| Cross-table filtering | [Model relationships](datamodel.md); load tables first |
| Interactive summary | [PivotTables](pivottable.md); model setup only when needed |
| A chart following PivotTable fields/filters | [Live PivotCharts](chart.md), not a chart of displayed cells |
| Requested report layout | [Optional formatting](report-formatting.md) |

## CLI batch jobs

MCP has no equivalent batch tool: await dependent calls. CLI batch avoids
repeated process startup for a known sequence. It is not a transaction.
Use individual commands when the next step needs inspection.

Commands still run one at a time through the same execution path as individual
CLI calls. Standard output contains only the JSON result lines and is flushed
after each command. Standard error shows when each command starts, completes,
or fails, including in quiet mode. A started line is not proof of completion.
If a wait expires, inspect the session's eventual state rather than
automatically repeating writes or assuming earlier changes were undone.

Batch commands use the same group and action names as the CLI
(`group.action`, for example `worksheetstyle.set-tab-color`). Arguments use
camel-case names, not CLI flags; for example, `--timeout-seconds` becomes
`timeoutSeconds`. Check every NDJSON success
flag, the exit status, and expected command count. Use `--stop-on-error`;
otherwise errors do not stop later commands. Keep save/close outside the batch.

### Identifying batch results

Both JSON-array input and one-command-per-line input return one JSON object per
output line. Each per-command object includes `commandSummary` alongside the
existing `index`, `command`, `success`, result, and error fields. For repeated
commands, summaries can distinguish their targets without repeating all input:

```text
range.set-values: Sales!A2:D21 (20 rows x 4 columns)
range.set-values: Forecast!B5:C8 (4 rows x 2 columns)
powerquery.refresh: queryName=SalesImport
chart.read: chartName=RevenueChart
diag.ping
```

Summaries always start with the command and use only these top-level input
details, in this fixed order:

| Input details | Summary |
|---|---|
| `sheetName`, `rangeAddress` | `Sales!A1:B2`; sheet names with spaces or punctuation are quoted |
| `sourceSheet`, `sourceRange` | `source=Sales!A1:B2` |
| `targetSheet`, `targetRange` | `target=Archive!C1:D2` |
| `tableName`, `queryName`, `pivotTableName`, `chartName` | Each supplied name as `key=value` |
| `oldName`, `newName` | Worksheet rename targets as `oldName=Before, newName=After` |
| `values` for `range.set-values`, `formulas` for `range.set-formulas` | Row/column counts for a nonempty rectangular inline array; never the cells themselves |

When only a sheet or range is supplied, it appears as `key=value`. An empty
sheet name is not replaced with a guessed active sheet. Unsupported details,
empty strings, wrong-type details, and text that cannot be decoded are omitted;
normal command validation still applies. With no supported details, the summary
is just the command.
File-based writes show available target details but do not read files to
calculate dimensions.

Each summary is at most 160 characters, including a trailing `...` if shortened.
Whitespace, control characters, and invisible formatting controls are replaced
with spaces; shortening keeps Unicode characters intact. Summaries are
display text, not unique identifiers or executable commands. Use `index` to
match a result to its original input.

The same summary rules apply to successful entries, validation errors, command
errors, and communication errors returned as per-command results.
`--stop-on-error` still stops after the first failed entry. A summary describes
the requested operation, not confirmation of completion or actual workbook
state. Batch-wide errors without per-command results remain unchanged.

Cell contents, formula text, VBA/M code, credentials, connection strings, free-form
messages, nested arguments, and file-path arguments are not copied into summaries.
Selected target names remain visible to the requesting user; summaries do not
scan or redact those names for secrets and are not safe to publish automatically.
Summaries are not sent to usage reporting. This field is CLI-only because MCP
has no batch tool.

### Running a batch job

This example assumes authorized edits and save/close of a workbook opened
exclusively for the job. `$path` is the supplied path, `Sales` already exists,
and target cells have been inspected. `commands.json` is a job-owned file.
For a reused session, report partial work without discarding earlier edits.

```cli
$path = [IO.Path]::GetFullPath($path)
$listed = excelcli -q session list | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) { throw "Cannot inspect existing sessions." }
if ($listed.sessions | Where-Object { $_.filePath -eq $path }) {
    throw "Reuse the existing session; this job must not discard its unsaved work."
}

@'
[
  {"command":"range.set-values","args":{"sheetName":"Sales","rangeAddress":"A1:B2","values":[["Product","Amount"],["Widget",1250]]}},
  {"command":"range.set-number-format","args":{"sheetName":"Sales","rangeAddress":"B2","formatCode":"$#,##0.00"}}
]
'@ | Set-Content commands.json -Encoding utf8

$opened = excelcli -q session open $path | ConvertFrom-Json
if ($LASTEXITCODE -ne 0 -or -not $opened.sessionId) { throw "Open failed; inspect session state." }
$sessionId = $opened.sessionId
$save = $false
$jobError = $null
try {
    $output = @(excelcli -q batch --session $sessionId --input commands.json --stop-on-error)
    $exit = $LASTEXITCODE
    $results = @($output | ConvertFrom-Json)
    if ($exit -ne 0 -or $results.Count -ne 2 -or ($results | Where-Object { -not $_.success })) {
        throw "Batch failed; this job will close without saving. Results: $($output -join ' ')"
    }
    $save = $true
}
catch {
    $jobError = $_
    throw
}
finally {
    try {
        $listed = excelcli -q session list | ConvertFrom-Json
        if ($LASTEXITCODE -ne 0) { throw "Cannot inspect cleanup state." }
        $owned = $listed.sessions | Where-Object { $_.sessionId -eq $sessionId }
        if (-not $owned -or -not $owned.canClose) { throw "Session missing or busy; inspect before shutdown." }
        $close = @('-q', 'session', 'close', '--session', $sessionId)
        if ($save) { $close += '--save' }
        excelcli @close
        if ($LASTEXITCODE -ne 0) { throw "Close failed; inspect the surviving session." }
    }
    catch {
        if ($jobError) {
            Write-Error "Cleanup also failed after $($jobError.Exception.Message): $_" -ErrorAction Continue
        }
        else { throw }
    }
}
```

Normal daemon shutdown leaves unsaved workbooks open and visible without saving
or discarding; it is not failure cleanup.
Report cleanup failures rather than killing Excel or claiming rollback.
See [recovery](behavioral-rules.md#sessions-and-failures).

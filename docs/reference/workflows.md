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

The executable provides the discovery path without a skill or external guide:

```cli
excelcli.exe --help
excelcli.exe batch --help
excelcli.exe range set-values --help
```

Batch help includes runnable PowerShell examples, input fields, session rules,
and result checks. Action-specific help maps its CLI flags to the exact
batch argument names. For example, `--sheet-name` maps to `sheetName`, and
`--range-address` maps to `rangeAddress`.
The entry's `sessionId` stays outside `args`. Arrays and objects are supplied
directly as JSON values, not strings containing JSON.

Both `--input <file>` and stdin accept a JSON array or one command per line.
Omitting `--input` or passing `--input -` reads stdin. For example, this
PowerShell check needs no workbook:

```cli
@'
[
  {"command":"diag.ping"},
  {"command":"diag.echo","args":{"message":"hello"}}
]
'@ | excelcli.exe batch --stop-on-error
```

An unrecognized space-separated command such as `diag ping` is rejected, with
the verified dotted spelling suggested when available. It is not automatically
corrected or executed. Invalid argument errors retain the accepted names and
point to action help; malformed input errors point to batch help.

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

### Reading results

Individual CLI operations and batch items use the same JSON object:

```json
{"command":"range.get-values","success":true,"durationMs":12,"data":{"sheetName":"Sales","rangeAddress":"$A$1:$B$1","rowCount":1,"columnCount":2,"values":[[false,0]],"cellErrors":[]}}
```

Batch items add a zero-based `index` that matches the original input. There is
no `commandSummary` or nested `result` wrapper. Business fields are under
`data`; an operation without returned data can omit it. Read `success` and the
process exit code: delivery alone does not establish that the operation worked.
Failures exit nonzero and put the explanation under `error.message`, with
`error.category`, the original cause and session details when available.
Warnings and `partialState` are outside `data` and cannot be hidden by choosing
fields. Cancellation does not undo earlier changes.

`durationMs` measures the operation through its returned outcome, including
validation and, for individual commands, connection/startup. It excludes writing
the output. Batch timing excludes shared startup and earlier items. A feature's
own `data.durationMs` remains separate because it can measure different work.

MCP keeps its existing business-object result rather than the CLI outer object;
its actual tool error flag reports failures, and text and structured results
agree.

Results keep every array position in its original order, including empty arrays. Warnings,
cell errors, skipped reasons, partial results, inspection limits, dimensions,
and continuation details remain visible. Exact cell reads never shorten a
requested grid. General result-field selection is not supported.

For a scalar read, such as a chart axis number-format string, the complete
value is under CLI `data` or MCP `result`.

MCP screenshots return the native
image with its actual format and captured bounds; CLI `--output` still writes
the complete image. Capture limits and warnings remain visible.

Reads return full business data, except `powerquery.list` omits formula previews.
Use `powerquery view` (MCP: `powerquery_read` action `view`) for full M code. Use existing
bounded reads such as `workbook.inspect` and existing paging options when less
data is needed.

Operation output is JSON even without `--quiet` or redirection. Diagnostics go
to standard error; help and version output remain human-readable. Scripts using
the previous output must change root business-field lookups to `data`, batch
`result` lookups to `data`, and root error-text lookups to `error.message`.

### Running a batch job

This example assumes authorized edits and save/close of a workbook opened
exclusively for the job. `$path` is the supplied path, `Sales` already exists,
and target cells have been inspected. `commands.json` is a job-owned file.
For a reused session, report partial work without discarding earlier edits.

```cli
$path = [IO.Path]::GetFullPath($path)
$listed = excelcli -q session list | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) { throw "Cannot inspect existing sessions." }
if ($listed.data.sessions | Where-Object { $_.filePath -eq $path }) {
    throw "Reuse the existing session; this job must not discard its unsaved work."
}

@'
[
  {"command":"range.set-values","args":{"sheetName":"Sales","rangeAddress":"A1:B2","values":[["Product","Amount"],["Widget",1250]]}},
  {"command":"range.set-number-format","args":{"sheetName":"Sales","rangeAddress":"B2","formatCode":"$#,##0.00"}}
]
'@ | Set-Content commands.json -Encoding utf8

$opened = excelcli -q session open $path | ConvertFrom-Json
if ($LASTEXITCODE -ne 0 -or -not $opened.data.sessionId) { throw "Open failed; inspect session state." }
$sessionId = $opened.data.sessionId
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
        $owned = $listed.data.sessions | Where-Object { $_.sessionId -eq $sessionId }
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

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

Batch commands use the same group and action names as the CLI
(`group.action`, for example `worksheetstyle.set-tab-color`). Arguments use
camel-case names, not CLI flags; for example, `--timeout-seconds` becomes
`timeoutSeconds`. Check every NDJSON success
flag, the exit status, and expected command count. Use `--stop-on-error`;
otherwise errors do not stop later commands. Keep save/close outside the batch.

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

Normal daemon shutdown may save remaining sessions; it is not failure cleanup.
Report cleanup failures rather than killing Excel or claiming rollback.
See [recovery](behavioral-rules.md#sessions-and-failures).

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

## Finishing the requested work

Use checks proportionate to the request, not a full-workbook audit. A single-cell
read needs only the requested answer; a simple edit needs confirmation of the
affected result, not unrelated screenshots or an extra Checks sheet.

1. **Confirm the requested result.** Read back changed cells or objects and
   compare them with the request. Check meaningful operation failures and
   warnings; a successful write only establishes that an operation ran. When
   several sheets were created or changed, inspect the actual sheet collection
   and intended deliverables so missing content or unfinished sheets are not
   overlooked.
2. **Check relevant calculations.** Inspect representative values, formulas,
   references, and returned cell-error details. Check that derived outputs use
   formulas rather than fixed numbers when required. Plausible numbers and
   an absence of formula errors do not prove the model is correct. Follow
   [calculation guidance](calculation.md) and existing authorization rules;
   a read-only check does not authorize recalculation or refresh.
3. **Check relationships that matter to the task.** Compare relevant totals
   with components or source data, and check units, signs, and assumptions.
   Do not change customer inputs or assumptions to test sensitivity; scenario
   changes require an explicit request, and tests that change inputs belong in
   synthetic fixtures. Financial-model-specific guidance is tracked separately
   in [#938](https://github.com/sbroenne/mcp-server-excel/issues/938). It is not
   yet available and is not required for this checklist.
4. **Check appearance when it matters.** Inspect displayed numbers and changed
   report/chart areas using the existing
   [screenshot guidance](screenshot.md) when an interactive desktop is available.
   Stored values and format codes alone do not establish what the user sees.
   Repair unreadable values, clipped labels, broken charts, and important
   overlaps locally. If capture is unavailable or incomplete, state the visual
   limitation; do not treat it as a passed check or block an otherwise authorized
   save/close solely for that reason.
5. **Save and close only as requested.** In-memory success is not a saved file.
   Save explicitly when requested, following the existing
   [workbook](workbook.md) and
   [session rules](behavioral-rules.md#sessions-and-failures). After a failed save,
   retain the session and report remaining unsaved changes or uncertainty.
   Check live readiness before an authorized close; an idle request count alone
   is insufficient. Save/reopen only when persistence is the behavior being
   checked, comparing the relevant reopened contents without changing inputs or
   unrelated state. Do not open extra workbook copies for routine verification.
6. **Deliver honestly.** Briefly state what was produced and any meaningful
   remaining limitation. Identify the material areas checked and any relevant
   unavailable checks; do not claim visual, calculation, or persistence checks
   that were not performed. A sample does not justify "all errors fixed" or
   "model clean." If the user reports an unchanged or wrong result, first inspect
   the disputed current state rather than relying on a prior successful call.

For large workbooks, use targeted reads, manageable blocks, and short summaries,
not a dump of every cell. Use existing bulk operations or the CLI batch below
when appropriate; await dependent MCP calls. Preserve workbook structure and
unrelated correct work. Stop once meaningful problems are repaired; repeat
inspection after a relevant change or new evidence, not as an identical check
or cosmetic redesign loop. Missing verification data is a limitation, not success.

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

This example runs authorized edits and leaves the session open for inspection.
`$path` is the supplied path, `Sales` already exists, and the empty target cells
have been inspected. `commands.json` is a job-owned file. The session check below
does not replace discovery of an already-open workbook; follow
[session discovery](behavioral-rules.md#sessions-and-failures) before opening.

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

$info = excelcli -q workbook get-info --session $sessionId | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) { throw "Cannot inspect workbook state; session $sessionId remains open." }
$info | Select-Object readOnly, autoSaveOn
if ($info.readOnly) { throw "Workbook is read-only; no batch edits requested." }

$output = @(excelcli -q batch --session $sessionId --input commands.json --stop-on-error)
$exit = $LASTEXITCODE
$results = @($output | ConvertFrom-Json)
if ($exit -ne 0 -or $results.Count -ne 2 -or ($results | Where-Object { -not $_.success })) {
    throw "Batch incomplete; inspect partial state in session $sessionId before saving or discarding. Results: $($output -join ' ')"
}
```

Successful batch results do not prove the requested result. Read back `Sales!A1:B2`
and the number format of `B2`; compare the values with the input and the format
with `$#,##0.00`. Apply the
[finishing checklist](#finishing-the-requested-work), including appearance checks
when relevant and available. Save and close separately only when authorized and
ready. On failure, keep the session available and inspect partial state before
deciding whether to save or discard; closing without saving cannot undo edits
already saved by AutoSave.

Normal daemon shutdown leaves unsaved workbooks open and visible without saving
or discarding; it is not failure cleanup.
Report cleanup failures rather than killing Excel or claiming rollback.
See [recovery](behavioral-rules.md#sessions-and-failures).

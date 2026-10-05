# ExcelMcp agent evaluations

**This entire Python suite is on-demand only, not part of the normal development
lifecycle**, including its offline, Excel fixture, SDK-discovery, and live-agent
checks. It is not required for ordinary product validation, commits, PRs, or
merges. Follow the
[repository's on-demand evaluation policy](../AGENTS.md#build-and-validation):
use the setup and commands below only for explicitly requested evaluation work.
Missing evaluation dependencies do not block product delivery; required .NET,
Excel, and normal Git-hook checks remain unchanged.

These manual tests ask real agents to use the MCP Server and `excelcli`. They
answer a different question from integration tests: can an agent discover the
right workflow, complete it correctly, and respect the user's permissions?
An agent saying "done" is not proof of a correct workbook.

The project uses **pytest-skill-engineering 1.0.3**, with native JSON evidence
schema 4.0. Each execution is one attempt. The removed retries, AI judges,
Azure report summaries, and HTML/Markdown report options are not used.

After rebasing onto main commit `792511e4` and adopting public 1.0.3, the
71 offline checks, four unbilled skill-discovery checks, and 19 distinct Excel
checker proofs passed (the five affected proofs were rerun after fixture fixes).
The completed public 1.0.3 real-world comparison then verified all 48 unique
cases. Skills were read in all 24 treatment cases, but neither entry point gained
correctness: recorded tokens increased **23.4% for MCP** and **72.4% for CLI**.
The resumed round reserved **50 of its approved 100 attempts**, including two
interruptions with unknown usage. The historical 1.0.2 results below are separate.
The current source now contains only narrowly scoped formatting skills; these
historical results describe the broad skills, not the new ones.

| Entry point | Without skill | With broad skill | Verified cases per condition |
|-------------|--------------:|-----------------:|-----------------------------:|
| MCP | 4,312,584 recorded tokens | 5,320,118 recorded tokens | 12/12 |
| CLI | 1,794,556 recorded tokens | 3,092,967 recorded tokens | 12/12 |

The [compact receipt](evidence/real-world-1.0.3.json) contains case-level metrics,
frozen skill/checker hashes, and invocation boundaries without private traces.
Interrupted usage is excluded from the percentages, not assumed zero. Results
cross resumed invocations and a recorder-code change. One capable model and two
repetitions cannot establish universal skill value, and equal success leaves no
correctness difference to measure.

Skill-read counts include only calls explicitly marked successful. Token
percentages require the same verified `(task, repetition)` cases in both
conditions, complete recorded usage, and a nonzero baseline. Otherwise the
receipt omits that percentage and explains why in `token_comparison_exclusions`.

## Setup

Live evaluations require Windows, desktop Excel, the SDK selected by
`global.json`, GitHub Copilot access, and an available model. Authenticate with
`gh auth login`, `GITHUB_TOKEN`, or `GH_TOKEN`. Azure OpenAI is not required.

From the repository root:

```powershell
dotnet build Sbroenne.ExcelMcp.sln -c Release
& .\scripts\Build-AgentSkills.ps1 -GenerateOnly
Set-Location llm-tests
uv sync
```

Generate skills from their canonical sources. Do not test stale installed
copies or edit generated `SKILL.md` files. The published-plugin experiment below
deliberately uses an installed published package instead of generated sources.

## Run the right checks

Checks of the harness and transport need neither Excel nor model access:

```powershell
uv run python -m unittest test_eval_harness.py test_cli_mcp_server.py test_cli_result_assertions.py test_consent_scenarios.py test_skill_value_checks.py test_skill_value_contracts.py test_spreadsheetbench.py test_aggregate_skill_value.py test_formatting_value.py -v
```

Checks of the independent workbook reader and experiment fixtures need Excel
but do not call a model:

```powershell
uv run python -m unittest test_workbook_assertions.py test_skill_value_preparation.py test_owned_workbook_cleanup.py -v
```

Real SDK skill-discovery checks require GitHub access but neither Excel nor
model messages:

```powershell
uv run pytest test_skill_discovery.py -v
```

Retained live scenarios exercise chart positioning, slicer outcomes, and
permission handling. Run only affected files. For example:

```powershell
uv run pytest mcp_tests\test_mcp_chart_positioning.py --aitest-json TestResults\chart-mcp-new-run.json -v
uv run pytest cli\test_cli_consent.py --aitest-json TestResults\consent-cli-new-run.json -v
```

Run every Excel-dependent command **sequentially**. Do not use parallel pytest
workers or overlap these runs with other Excel tests. The harness owns private
CLI pipes and isolated working directories; it must not stop a user's default
service. GitHub prerequisites apply to live tests, not offline checks.

After outcome verification, teardown closes only the fixture's exact workbook
path, without saving. This includes MCP-owned workbooks, whose sessions are not
visible through the private CLI pipe. The cleanup proof checks that a second
workbook remains open and that unsaved changes are discarded. Cleanup errors
must remain visible; never kill Excel by process name or stop a default service.

`EXCEL_MCP_SERVER_COMMAND` and `EXCEL_CLI_COMMAND` can override the executable.
Ordinary evaluations default to `gpt-6.1-sol`; `EXCEL_LLM_MODEL` can explicitly
select another available model. Skill-value comparisons are always pinned to
`gpt-6.1-sol` and never silently fall back to `auto`.

### Published CLI plugin discovery

`test_published_cli_discovery.py` is a focused, opt-in experiment: does an
ordinary workbook request activate the published `excel-cli` skill and use
the actual packaged `bin\start-cli.ps1` launcher? It loads both installed
skills (`excel-cli` and `excel-cli-report-formatting`) with native tools and
generic instructions, not the evaluation harness's CLI MCP bridge or custom
workspace tool. The request gives no CLI, skill, or launcher tutorial.

Install the published `excel-cli` plugin using your client's plugin manager
from [the published plugin repository](https://github.com/sbroenne/mcp-server-excel-plugins/tree/main/plugins/excel-cli).
Select that installed plugin's root directory, containing `plugin.json`,
`bin\start-cli.ps1`, and both `skills` directories. Treat it as immutable input;
record its release/revision separately. This checks discovery and use of that
package, **not installation itself**, MCP skill selection, a skill-value
comparison, or universal reliability. The current harness's `CopilotEval.from_plugin`
does not accept the published author-object metadata, so the experiment loads
the actual individual skill directories without modifying the plugin.

From `llm-tests`, with the locked dependencies installed (`uv sync --locked`):

```powershell
# Unpaid checker proofs: no plugin, Excel, or model access needed.
uv run --locked pytest test_published_cli_discovery.py -m offline -v

# Replace this with your installed published plugin's root directory.
$env:EXCEL_PUBLISHED_CLI_PLUGIN = 'C:\path\to\installed\excel-cli'
# Unpaid SDK preflight: validates both SKILL.md sources and enabled skills,
# checks native powershell/skill tools, and forbids any model message.
uv run --locked pytest test_published_cli_discovery.py::test_published_cli_unpaid_preflight -v

# Only after those pass and one paid attempt is approved:
$env:EXCEL_RUN_PUBLISHED_CLI_DISCOVERY = '1'
uv run --locked pytest test_published_cli_discovery.py::test_live_published_cli_discovery --aitest-json TestResults\published-cli-new-run.json -v
Remove-Item Env:\EXCEL_RUN_PUBLISHED_CLI_DISCOVERY
Remove-Item Env:\EXCEL_PUBLISHED_CLI_PLUGIN
```

Do not add iterations or silently retry. Normal collection without the live
opt-in skips the paid test. Missing prerequisites may skip; a supplied invalid
plugin or incorrect result fails. The live attempt repeats the unpaid preflight,
uses the shared model/budgets, a fresh pytest task directory, and a unique private
CLI pipe. Run it sequentially with all other Excel-dependent work.

Checks require successful correlated tool completions, activation of `excel-cli`,
successful packaged-launcher value/formula writes and save-and-close, no open
owned workbook before teardown, and independent Excel recalculation of saved
`A1:D5` values/formulas. The quoted product name must survive; line totals must
be 37.5, 350.5, and 35, with grand total 423. Negative proofs reject missing/wrong
sheets, incorrect values, constants or wrong formulas, help-only/global-CLI
evidence, failed/incomplete calls, missing skill activation, and missing close.
Cleanup closes only the exact synthetic workbook and stops only the private
service; errors remain visible. Native JSON and the temporary COM snapshot
contain local paths and must stay private, not be committed or posted publicly.

## Measure whether skills help

`test_skill_value.py` defaults to the **formatting** suite: two presentation tasks
with two repetitions, plus a read-only audit and model refresh with one repetition,
each through both entry points with and without the matching formatting skill.
That is **24 paid executions**: 16 positive-trigger and 8 non-trigger cases.
It is skipped unless explicitly enabled:

```powershell
uv run pytest test_skill_value.py --collect-only -q
uv run pytest test_skill_value.py --run-skill-value --skill-value-output TestResults\skill-value-new-run --aitest-json TestResults\skill-value-new-run-native.json -v
```

Choose a **new** evidence directory and JSON filename for every invocation.
Existing comparison directories are rejected rather than mistaken for current
evidence. Do not add `--aitest-iterations`: the repetitions are already
part of this matrix, and that option repeats every collected test.
Every smoke, retry, and rerun counts toward the approved paid-execution limit.
When previous attempts have already used part of the allowance, record them
with `--skill-value-prior-attempts` and choose a balanced smaller matrix with
`--skill-value-repetitions 1` or `2`. Planned plus prior attempts must fit
`--skill-value-ceiling`, which defaults to 60. Use 100 only with explicit
authorization. The harness checks this before sending any model message.
One formatting repetition collects 16 cases. The older five-task
`--skill-value-suite business` defaults to three repetitions (60 cases);
two collect 40. Smaller samples reduce confidence.

Presentation checks preserve values, formulas, identifiers, analysis objects,
notes, and calculation mode; check USD and fractional percentage formats,
bold headers, and rendered numbers without `####`; and test financial-model
input/formula colours, parentheses for negatives, and dashes for zeros.
These checks do not prove that every header is visually readable or judge
arbitrary aesthetics. Skill selection is saved separately from workbook
correctness: positive tasks should read the skill, ordinary tasks should not.
A correct workbook does not hide a wrong selection result. This new suite has
not yet been run with paid agents, so there is no claimed formatting benefit.

The CLI evaluation wrapper uses the MCP 2 server API selected by `uv.lock`.
Its transport smoke test, CLI call-recording checks, and consent-assertion
regressions need neither Excel nor model access:

```powershell
uv run python -m unittest test_cli_mcp_server.py test_cli_result_assertions.py test_consent_scenarios.py -v
```

CLI assertions recognize both `excel_execute` and `excel-cli-excel_execute`.
Legacy SDK outputs recorded only in tool turns can still be decoded for
inspection. They cannot establish a completed execution in current evaluations:
each call also needs the package's correlated completion evidence. Missing or
conflicting records fail the evaluation rather than hiding command failures.

| Task | Independently checked result |
|---|---|
| Simple report | Real named Excel Table, requested data, saved total formulas, correct totals, two-decimal data and total formats |
| Bulk targeted updates | 60 prices changed in 120 existing rows; other values, formulas, formats, named table/chart, and original calculation mode preserved; chart series follow the updated values |
| Power Query recovery | Existing failed query repaired; typed dates/numbers loaded at the right destination; note and similarly named query preserved; no duplicates |
| Model-backed refresh | New orders reach the existing model/PivotTable/PivotChart; correct category totals; relationship, measure, tabular layout, and linkage preserved |
| Read-only audit | Correct discrepancy reported; no workbook-changing calls or saved-file changes; CLI's existing unsaved session, note, hidden window, and manual mode preserved |

### Real-world cases and native Excel workflows

The opt-in `--skill-value-suite real-world` comparison combines four public
workbook problems with the existing Power Query and Data Model/PowerPivot
workflows. SpreadsheetBench's cell-answer checks alone do not establish that
an agent can repair queries or refresh a genuinely model-backed report.

| Task | Source | Independently checked result |
|---|---|---|
| Grouped transaction totals | SpreadsheetBench Verified `13-1` | Date/reference groups, sorted results, section totals, and untouched source data |
| Monthly inventory comparison | SpreadsheetBench Verified `267-21` | Correct ID matches across two sheets and `-` for missing matches |
| Last-record deduplication | SpreadsheetBench Verified `280-17` | Last occurrence retained for each key, correct order, no leftover rows |
| Date/text conditional totals | SpreadsheetBench Verified `38823` | Every requested total remains a formula; changing units and a date boundary produces independently calculated new totals |
| Power Query recovery | Authored business fixture | Failed query repaired, typed data loaded at the requested location, similarly named query and notes preserved |
| Data Model/PowerPivot refresh | Authored business fixture | New orders reach the existing model, DAX measure, model-backed PivotTable and linked PivotChart; relationship and named objects preserved |

PowerPivot here means Excel's Data Model, DAX, and model-backed analysis, not
automation of the separate PowerPivot window. These two authored tasks cover
repair and refresh, not designing arbitrary models from scratch.

The public inputs come from
[RUCKBReasoning/SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench),
pinned to revision `49b73a94775fb489063f60ca1865e3a650079a79`.
The Verified 400 archive has SHA256
`10ef893dd29cb13ab97143ea787e68cdc9574a13873ab9a54e50b31dc03fc949`.
The upstream project declares
[CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/).
Keep that attribution, identify modifications, and apply the same licence to
redistributed dataset content or adapted benchmark requests. The dataset
licence is separate from the harness code's licence.

Download into the ignored cache and check the fixtures without any agent calls:

```powershell
uv run python spreadsheetbench.py fetch
uv run python spreadsheetbench.py manifest
uv run python -m unittest test_spreadsheetbench.py -v
uv run python -m unittest test_spreadsheetbench_workbooks.py -v
uv run pytest test_skill_value.py --skill-value-suite real-world --collect-only -q
```

Only `test_spreadsheetbench_workbooks.py` needs Excel; collection and the offline loader
checks do not. The workbook checks verify the official answer regions, reject
unchanged inputs and incorrect answers, and reject constant formulas that only
match the original totals. The data-change probe operates inside a read-only
Excel instance and closes without saving; it must not alter the saved workbook.
Full-sheet inspection is bounded to 10,000 used cells per worksheet; exceeding
that limit fails the case explicitly rather than silently checking only a corner.
Answers and expected values are never placed in the agent's task directory.
No third-party workbooks or full instructions are committed.

Fixture setup and deliberate checker mutations use CLI `--overwrite-policy allow`
for test-owned cells. This is not injected into agent requests or tools: agents
must discover the current overwrite policy from normal help/schemas and decide
whether the user's requested replacement permits it. The batch-example proof
reads the current shared workflow guide and checks both complete-save and
partial-failure discard behavior.

The recorder configures an adjacent `.events.jsonl` file through the public
`CopilotEval.extra_config["on_event"]` callback, but the completed real comparison
received **zero live events**. Mocked callback checks do not prove live streaming;
the callback's actual SDK delivery remains unverified. Atomic attempt snapshots
and saving returned execution before workbook verification did work.
Caught cancellation marks the
attempt interrupted; an abruptly killed process may still leave `running` or
`verifying`, which must be treated as incomplete, not as a verified failure or
zero usage. Keep these local event journals private along with other raw traces.
Do not change frozen checkers during a run. A resumed invocation uses a new
evidence directory, keeps all earlier attempts, and records their reserved count
with `--skill-value-prior-attempts`.

This is an **adapted pilot**, not an official SpreadsheetBench score. Requests
are explicitly limited to the current workbook, retaining formula requirements,
and add preservation and save/close requirements. Existing worked examples in
the public inputs are retained. Official answer cells provide the reference;
incidental changes to header number formats in answer files are not adopted.
Original source values, formulas, number formats, sheet order,
existing analysis objects, and calculation mode remain protected. Case `66-24`
was considered but excluded because its answer file changes unrelated source
content. Each selected instruction has one published input variant; repetitions
reuse its contents, not independent datasets. Public cases may also be known to
the model, so success alone does not prove generalisation.

This suite defaults to two repetitions: six tasks, two entry points, two skill
conditions, **48 paid executions**. One repetition collects 24; three collect
72 and exceed the default 60-attempt ceiling unless a smaller balanced subset
is selected or a higher ceiling is explicitly authorized. The business suite
remains available at 60 but is no longer the default. All suites
use the same isolation, discovery, evidence, usage, and budget checks.

The real-world suite completed with the results at the top of this document.
The original 60-attempt round remains separate from the resumed 100-attempt
allowance. No new formatting-skill or example comparison is authorized merely
by collecting tests or preparing fixtures. Competition downloads are not included:
free availability does
not establish permission to redistribute or use them in this public test suite.

### What is held constant

Both conditions get the same model, request, fixture contents, role instructions,
tool access, permissions, **600-second limit**, and **80 admitted tool calls**.
`max_turns` is advisory, not a hard limit. Each execution gets a fresh task
folder and owned Excel sessions. Condition order alternates between repetitions.
The only intended difference within each entry point is whether the matching
generated skill is available.

The SDK's `empty` mode disables ambient configuration, file hooks, and
on-demand instruction discovery. It also defaults skill loading off, so both
builders explicitly set the public SDK option `enable_skills=True` while
keeping ambient discovery disabled. Before any paid comparison, fresh isolated
SDK sessions must discover exactly the matching skill in each treatment and no
skills in either baseline. Configuring a directory is not proof of availability.
The discovered metadata is recorded in the manifest, including resolved source
paths. First, the probes call public `pytest_skill_engineering.load_skill` on
each actual supplied skill directory, exercising the same local validation as
the runner. Then they check real SDK discovery with both public `send` and
`send_and_wait` forbidden. This catches package validation failures that SDK-only
discovery misses, without sending a model message or starting an Excel server.
Shell/process tools and arbitrary file access
are excluded. The shared `workspace` tool can read task files and the supplied
skill's references, and write text/JSON/CSV/M inputs inside the task directory.
It cannot read repository code, checkers, expected answers, or arbitrary paths,
or modify a workbook directly. Both conditions retain normal MCP descriptions,
server instructions, and CLI `--help`.

Do not require the treatment agent to read the skill. Record whether it actually
uses `skill`, reads `SKILL.md`, or reads references: an available but unused skill has not delivered
practical value for that execution. Source/skill/checker hashes are frozen in
the run manifest. Do not tune skills during a baseline comparison.

### What counts as benefit

Primary evidence is **verified correctness and safety**, not final prose,
particular tool choices, or use of a screenshot. A skill earns a positive signal
if it improves verified completion, or reduces measured usage by **at least
20% without worse correctness or safety**. Failed attempts and skill-reading
overhead count in usage per verified completion. Batch subcommands are recorded
separately from outer agent calls.

Use recorded premium requests when available; otherwise use complete recorded
input-plus-output tokens. Missing usage/pricing is **unknown**, not zero.
The SDK currently has a known premium-request capture gap; do not interpret a
default zero in a native report as free execution. Tokens measure model usage,
not dollar cost, and no unsupported pricing is invented.

Keep task failures, budget exhaustion, execution/setup failures, and incomplete
capture distinct. Invalid execution/capture or a harness exception stops the
comparison and persists its failure record, including any returned execution.
Each
condition has its own pytest outcome and verification record; a paired helper's
shared outcome is not used as evidence that both sides succeeded.

Three repetitions on one model cannot establish universal benefit or
uselessness. Report raw paired results, inconsistent differences, unused skills,
and uncertainty. Small differences are inconclusive. Do not remove shipped
skills merely because this small experiment finds no clear benefit.

### Measured results: public 1.0.2

**Neither skill earned its overhead in these tasks on `gpt-6.1-sol`.** All
sixteen matched pairs produced independently verified results, with no observed
safety/preservation failures. The treatment read guidance in every execution;
this is not a comparison against missing or ignored skills. Every matched
treatment used more recorded input-plus-output tokens than its baseline.

The corrected, balanced comparison contains eight executions per condition
and entry point: control/audit once, and each demanding task twice.

| Entry point | Verified without / with skill | Tokens without skill | Tokens with skill | Additional tokens | Agent calls without / with skill |
|---|---|---:|---:|---:|---|
| MCP | 8/8 / 8/8 | 2,653,719 | 3,475,767 | **31.0%** | 146 / 172 |
| CLI | 8/8 / 8/8 | 1,062,095 | 1,894,770 | **78.4%** | 179 / 153 |

CLI guidance reduced outer agent calls by 14.5%, partly through batching
(87 executed batch subcommands), but **fewer calls did not mean lower model
usage**. MCP treatment made more calls. Agent-only duration was 9.0% higher
for MCP and 26.3% higher for CLI; this excludes fixture preparation and
independent inspection and is not a billing measure.

| Task | Pairs per entry point | Additional MCP tokens | Additional CLI tokens |
|---|---:|---:|---:|
| Simple report | 1 | 36.6% | 49.6% |
| Bulk targeted updates | 2 | 22.8% | 122.0% |
| Power Query recovery | 2 | 31.6% | 91.4% |
| Model-backed refresh | 2 | 36.1% | 46.6% |
| Read-only audit | 1 | 28.1% | 33.6% |

These are ratios of summed task tokens, not averages of percentages. All paired
outcomes passed; no pair met the predeclared 20% usage-reduction threshold.
The MCP skill had sixteen successful guidance reads in the balanced comparison,
and the CLI skill had forty-two; both baselines had zero.

**Recommendation:** keep the outcome-checked LLM scenarios, but trim/reconsider
both skills rather than claim a proven reliability or efficiency gain.
Prioritize concise, non-obvious ExcelMcp rules over repeated help/tutorials.
CLI runs also exposed ordinary-command versus batch-command naming mistakes.
Do not require the full paid matrix on every change: run affected scenarios,
and repeat a value comparison when guidance materially changes. Shipped skill
content was not changed or removed during this experiment.

This does **not** prove that skills are universally useless. Baselines already
completed every task, leaving no correctness improvement to observe. Only one
strong model and one/two repetitions were used. There was no comparison of
shorter skills, individual references, other models, or harder unfamiliar tasks.
CLI was exercised through the existing test wrapper, not a native shell;
recovered quoting and batch errors remain in the usage totals. Token totals
include repeated model input and are not dollar costs. Premium-request billing
was unavailable, so complete tokens are the declared fallback.

#### Corrections and spending

The 32-case run initially reported 31 passes, one scoring failure, and two
teardown errors. The failed CLI bulk treatment encountered a rejected
`batch --commands` invocation before eventually saving/closing. The checker
incorrectly required a batch input for that proven parser rejection and never
reopened its final workbook. That original execution remains **unverified**,
not an agent failure or a silently promoted success. A fresh treatment with
the same request, workbook checks, model, and skill passed and replaces only
that invalid scoring slot in the balanced tables above.

Both original MCP audit outcomes had passed before cleanup failed: MCP and CLI
do not share sessions. Cleanup now locates only the exact owned workbook through
Excel COM and discards its changes after verification. An unbilled live check
proved that another workbook stays open; the two audits were also repeated
successfully, with no teardown errors. That supplemental MCP pair used 185,410
baseline tokens and 237,621 treatment tokens (28.2% extra), consistent with the
original observation. It is not an extra repetition for every task.

No execution was erased. Across **all 35 public-1.0.2 executions**, including
the unverified bulk run and the supplemental audits:

| Entry point / condition | Verified / executed | Total tokens | Tokens per verified completion |
|---|---|---:|---:|
| MCP without skill | 9/9 | 2,839,129 | 315,459 |
| MCP with skill | 9/9 | 3,713,388 | 412,599 |
| CLI without skill | 8/8 | 1,062,095 | 132,762 |
| CLI with skill | 8/9 | 2,162,207 | 270,276 |

Including the original 267,437-token scoring-invalid execution makes CLI
treatment usage per verified completion **103.6% higher**, not more efficient.
The balanced result excludes its unverified outcome; this all-attempt view
keeps its overhead visible. These two summaries have explicitly different
denominators.

Earlier package/setup diagnostics are not pooled into these results. Initial
skill discovery was invalidated; upstream
[sbroenne/pytest-skill-engineering#110](https://github.com/sbroenne/pytest-skill-engineering/pull/110)
fixed explicit loading in 1.0.1, and
[sbroenne/pytest-skill-engineering#112](https://github.com/sbroenne/pytest-skill-engineering/pull/112)
fixed nested CLI references in the adopted public 1.0.2. Local duplicate SDK
output was corrected to a text-only result with strict agreement checks.
Twenty-five earlier allocated slots were conservatively reserved, including
a known pre-model setup rejection; 32 main cases plus three corrections bring
the ledger to **60/60 reserved slots**. This is not a claim of sixty measured
or billed executions. Further paid work needs a new approved allowance.

Exact completed comparison commands, from `llm-tests`:

```powershell
uv run pytest test_skill_value.py --run-skill-value --skill-value-repetitions 2 --skill-value-prior-attempts 25 --skill-value-output TestResults\skill-value-release-1.0.2 --aitest-json TestResults\skill-value-release-1.0.2-native.json -k 'not ((control or read-only-audit) and rep2)' -v
# Initial result: 31 passed, 1 scoring failure, 2 teardown errors; all 32 cases recorded.
uv run pytest test_skill_value.py --run-skill-value --skill-value-repetitions 1 --skill-value-prior-attempts 57 --skill-value-output TestResults\skill-value-release-1.0.2-corrections --aitest-json TestResults\skill-value-release-1.0.2-corrections-native.json -k 'bulk-update-cli-with-skill-rep1 or read-only-audit-mcp' -v
# Correction result: 3 passed, no teardown errors.
```

Both manifests confirm unchanged generated skills, canonical templates/shared
guides, task requests, and workbook criteria. Only batch-evidence accounting,
owned-session cleanup, and recording the cleanup helper's hash changed.
Subsequent offline hardening rejects unknown parser exit codes, preserves
indexed failed batch steps, and classifies missing batch input as capture
failure rather than an agent mistake. Stored-call accounting for all 35 release
executions was checked again; the original scoring-invalid workbook was not
reclassified as verified.
Do not rerun these commands against the exhausted allowance or reuse their
evidence paths.

### Evidence

`--aitest-json` writes native traces, configuration, tool outcomes, and usage.
`--skill-value-output` writes `manifest.json`, one verification record per
execution, and `summary.json`. Verification records also appear in native test
properties. The saved workbook is reopened by **separate Excel COM inspection**,
read-only, with macros/events disabled and links not updated; it is closed
without saving. The reader captures formulas, formats, query identity, model
relationships/measures, real PivotChart linkage, and saved values.

Raw evidence belongs in gitignored `TestResults`, not tracked reports. It can
contain local paths and synthetic workbook contents. Only sanitized measured
findings belong in public documentation. Reading saved evidence requires no
authentication or additional model calls.

### Migration validation

The final setup passed these focused checks. Commands below are from `llm-tests`
unless marked otherwise; none of the supporting checks sent model messages.

| Command | Result |
|---|---|
| `uv run python -m unittest test_eval_harness test_cli_mcp_server test_consent_scenarios test_skill_value_checks -v` | 41 passed |
| `uv run python -m unittest test_workbook_assertions test_skill_value_preparation test_owned_workbook_cleanup -v` | 14 passed with desktop Excel |
| `uv run pytest test_skill_discovery.py -v` | 4 real SDK preflights passed; sends forbidden |
| `uv run pytest --collect-only -q` | 144 cases collected; no executions |
| Root: `dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore --verbosity quiet` | Passed, zero warnings/errors |
| Root: `& .\scripts\Invoke-ExcelFreeTests.ps1 -Local -Contracts` | 147 passed across Core, CLI, and MCP |
| Root: `& .\scripts\check-doc-counts.ps1 -SkipBuild` | Passed after the Release build |
| Website: `.\.venv\Scripts\python.exe -m mkdocs build --strict --clean` | Passed |
| Website: `.\.venv\Scripts\python.exe audit_site.py` | 53 pages passed |
| Website: `.\.venv\Scripts\python.exe check_deploy_paths.py` | Passed |

Existing root checks `check-com-leaks.ps1`, `check-success-flag.ps1`,
`check-dynamic-casts.ps1`, and `check-workbook-package-access.ps1` all passed,
as did `git diff --check`. Product `scripts\Test-E2E.ps1` was **not run**:
this migration changes evaluations/docs, not product runtime code.

## Which previous tests were worth keeping?

Keep the chart-positioning, slicer, and permission scenarios: they check real
geometry, actual filters/totals, or authorized versus unauthorized changes.
The CLI-only pre-existing unsaved-session audit is also valuable. These are
agent-behavior checks, not duplicated API success tests.

The following old test functions were removed or replaced. In paired names,
`{cli,mcp}` means the corresponding function under each entry point. Service
test filenames below live under `tests/ExcelMcp.Service.Tests`; they remain
deterministic coverage, not claims that an agent can discover an operation.

| Previous functions | Decision and concrete replacement |
|---|---|
| `test_{cli,mcp}_file_and_worksheet_workflow` | Remove execution-success-only checks. `control` checks a saved workbook; `PersistentServiceSheetTests.Lifecycle.cs` verifies create, rename, delete, and copy against listed sheets. |
| `test_{cli,mcp}_range_set_get` | Remove duplicate smoke. `control` checks actual saved data; `PersistentServiceRangeValuesTests.Values.cs` covers typed round trips, dates, and wide ranges. |
| `test_{cli,mcp}_range_error_handling` | Remove misleading error test: the supposedly invalid empty range is valid. Real invalid/merged-range regressions remain in Service; do not preserve a non-error as recovery evidence. |
| `test_{cli,mcp}_table_create_query`, `test_{cli,mcp}_table_lifecycle` | Replace prose/success checks with `control` and `bulk-update`; create/delete/rename/resize/append/filter/totals remain in `PersistentServiceTablePreflightTests.Behavior.cs`. |
| `test_{cli,mcp}_range_updates`, `test_{cli,mcp}_table_updates` | Replace leaked expected answers and broad-word matches with saved formulas/values/formats in `bulk-update`, plus retained consent cases. |
| `test_{cli,mcp}_chart_updates` | Remove prose proxy. `bulk-update` checks preserved chart objects/geometry; `SetTitle_ValidTitle_SetsChartTitle` in `PersistentServiceChartFormattingTests.Appearance.cs` reads the actual title. |
| `test_{cli,mcp}_sheet_structural_changes` | Remove a narrated sequence with weak final-word checks. Real create/rename/delete/copy and range-edit behavior remain in Service. |
| `test_cli_calculation_mode_batch_flow`, `test_mcp_calculation_mode_batch_with_skill`, `test_mcp_calculation_mode_batch_no_skill` | Replace coached/manual-reset and prose checks with `bulk-update`, which checks correct formulas and restoration of the **original** mode, including manual. `PersistentServiceCalculationTests.cs` retains API scope/validation coverage. |
| `test_{cli,mcp}_chart_workflows` | Remove non-equivalent four/five-chart requests. Retained positioning scenarios check actual saved type/series/categories/geometry; Service chart lifecycle, data-source, and appearance tests retain API coverage. |
| `test_mcp_auto_position_no_skill`, `test_mcp_targetrange_no_skill` | Remove call/prose proxies, including explicit parameter coaching. Retained saved-position scenarios check the outcome, not whether the agent echoes a position. |
| `test_mcp_multi_chart_collision_no_skill`, `test_mcp_collision_warning_reaction_no_skill` | Remove unverified overlap claims. Existing real overlap rejection remains in `test_workbook_assertions.py`; a future distinct dashboard experiment must verify every shape, not infer quality from a warning/tool call. |
| `test_mcp_dashboard_layout_variants` | Remove screenshot-use tautology: the request told the agent to take a screenshot, and the test checked that it did. It did not measure layout improvement. |
| `test_{cli,mcp}_sales_report_workflow`, `test_{cli,mcp}_financial_report_automation` | Replace prompts leaking totals and differing cell placement with `control` and `model-refresh`; real formula/cross-sheet behavior remains in `PersistentServiceRangeFormulaTests.Cases.cs`. |
| `test_{cli,mcp}_pivottable_tabular_layout`, `test_{cli,mcp}_pivottable_compact_layout`, `test_{cli,mcp}_pivottable_outline_layout` | Replace mismatched requests and exact-name call-check errors with saved tabular model analysis in `model-refresh`; all three API layouts remain in `PersistentServicePivotTableTests.LayoutSubtotals.cs`. |
| `test_{cli,mcp}_star_schema_workflow`, `test_cli_powerquery_products_workflow`, `test_mcp_powerquery_amazon_workflow` | Replace exact-flag tutorials/obsolete M flags with `query-recovery` and `model-refresh`; `PersistentServicePowerQueryExactIdentityTests.cs` retains exact-identity and evaluation cleanup regressions. |
| `test_{cli,mcp}_styling_table_style`, `test_{cli,mcp}_styling_semantic_status`, `test_{cli,mcp}_styling_header_fill` | Remove checks for words such as "blue". `control` checks saved number formats and `bulk-update` checks preservation. Named styles remain covered by `PersistentServiceRangeSetStyleTests.Cases.cs`; no claim is made that generic word matches verified fills. |

This deliberately retires **24 weak live-test files**, not the entire LLM suite.
The obsolete failure-extraction script and tracked old-schema report were also
removed; neither represented current, trustworthy results.

## Writing evaluations

Write as an Excel user who knows the desired result but not ExcelMcp syntax.
Do not put command names, numeric enum values, exact flags, expected totals, or
recovery tutorials into a request just to make the agent pass. Fix unclear
guidance in its [canonical source](../docs/AGENT-SKILLS.md#authoring-and-packaging).

Use `build_excel_cli_eval`/`build_excel_mcp_eval`. Share equivalent requests and
independent outcome checks across entry points. Normal evaluations include the
matching skill; explicitly named tool-description-only experiments may omit it.
Assert completed successful calls where execution matters, not just attempted
arguments. Test checkers against deliberately wrong saved state before spending
on agents.

Investigate authentication, Excel, capture, and budget problems separately from
agent/product failures. Do not use skip/xfail to hide a product defect. Add a
focused deterministic regression before fixing an implementation bug. Update an
evaluation only when its request, setup, or checker is actually wrong.

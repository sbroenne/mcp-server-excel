# ExcelMcp LLM Integration Tests

LLM-powered integration tests for both ExcelMcp MCP Server and Excel CLI using pytest-skill-engineering.

**Platform scope:** the existing workbook/COM evaluations below target Windows.
They do not establish support for the experimental Apple Silicon macOS beta or
its gated features; use the separate [Mac acceptance requirements](../specs/MACOS-SUPPORT.md).

## Prerequisites

- Windows desktop with Microsoft Excel installed
- The .NET SDK selected by `global.json`
- GitHub Copilot access and authentication
- Azure OpenAI endpoint for the optional AI report summary (enabled by default)
- ExcelMcp MCP Server and CLI built/installed

### Azure OpenAI

Set the endpoint for Entra ID auth:

```powershell
$env:AZURE_OPENAI_ENDPOINT = "https://<your-resource>.openai.azure.com/"
```

## Setup

From this directory:

```powershell
uv sync
```

This installs the locked test dependencies, including pytest-skill-engineering
and its Copilot SDK dependency.

## Build MCP Server (Required)

```powershell
dotnet build ..\Sbroenne.ExcelMcp.sln -c Release
& ..\scripts\Build-AgentSkills.ps1 -GenerateOnly
```

## Run Tests (Manual Only)

### MCP Server tests

```powershell
uv run pytest -m mcp -v
```

### CLI tests

```powershell
uv run pytest -m cli -v
```

### All LLM tests

```powershell
uv run pytest -m aitest -v
```

## Configuration Overrides

The agent evaluations use GitHub Copilot, not Azure OpenAI. The default report
summary uses Azure OpenAI separately. To run without that optional summary:

```powershell
uv run pytest -o addopts= mcp_tests\test_mcp_chart_positioning.py -v
```

Run Excel-dependent commands sequentially. Never use parallel pytest workers
for these evaluations or overlap them with other Excel test runs.

The CLI evaluation wrapper uses the MCP 2 server API selected by `uv.lock`.
Its transport smoke test, CLI call-recording checks, and consent-assertion
regressions need neither Excel nor model access:

```powershell
uv run python -m unittest test_cli_mcp_server.py test_cli_result_assertions.py test_consent_scenarios.py -v
```

CLI assertions recognize both `excel_execute` and `excel-cli-excel_execute`.
When the SDK records command outputs only in tool turns, the assertions read
those outputs in order and require a result for every recorded CLI call.
Missing results fail the evaluation rather than hiding command failures.

- `EXCEL_MCP_SERVER_COMMAND` — override MCP server command (full command line)
- `EXCEL_CLI_COMMAND` — override CLI command (default: `excelcli`)
- `EXCEL_LLM_MODEL` — supported Copilot model ID; defaults to `auto` rather than
  pinning a retired model. Pin an available model for comparable evaluation runs.

Example:

```powershell
$env:EXCEL_MCP_SERVER_COMMAND = "d:\\source\\mcp-server-excel\\src\\ExcelMcp.McpServer\\bin\\Release\\net10.0\\Sbroenne.ExcelMcp.McpServer.exe"
$env:EXCEL_CLI_COMMAND = "excelcli"
```

## GitHub Copilot Authentication

These tests use the Copilot-backed `CopilotEval` harness from `pytest-skill-engineering`, so you must be authenticated with GitHub:

```powershell
gh auth login
```

Or set `GITHUB_TOKEN` in the environment before running `pytest`.

## Test Structure

- `test_mcp_*.py` — MCP Server workflows
- `test_cli_*.py` — CLI workflows
- `test_*calculation_mode*.py` — new calculation mode scenarios
- `test_*consent*.py` — paired clarification, read-only audit, visibility, and
  workbook-text permission scenarios, checked through calls and workbook state.
  The CLI also checks a pre-existing unsaved session in its private daemon.
- `Fixtures/` — shared test inputs (CSV/JSON/M files)
- `TestResults/` — HTML reports and artifacts

## Writing evaluations

These tests measure whether users' agents can discover a workflow through the
skill, CLI help, MCP descriptions, and error messages. They complement rather
than replace deterministic tests of Excel behavior.

### Requests and outcomes

Write as an Excel user who knows the desired result but not ExcelMcp syntax.
For example:

```text
Create a PivotTable showing Department and Team as rows and total Hours as
values. Use Compact layout, save the workbook, and report the totals.
```

Do not add the command name, numeric enum value, or a hint to run a particular
help command just to make the scenario pass. If such guidance is needed, put it
in the product's skill, help, description, or error message.

Prefer one coherent workflow per prompt. Split unrelated work into independent
tests instead of relying on a long chain of conversation state. Assert the
requested workbook outcome, tool execution, and relevant results; the agent
claiming success is not proof of a correct file. Avoid incidental phrasing
checks, but retain exact values when they are the outcome being tested.

### Harness and paired scenarios

Use `build_excel_cli_eval` or `build_excel_mcp_eval` from `conftest.py`, with the
shared default model, turn budget, timeout, and role instructions. Override a
budget only when a scenario needs it. Do not insert scenario-specific command
tutorials into the harness.

Normal skill evaluations include the matching skill directory. Tests
deliberately measuring tool descriptions without skills may omit it; identify
that purpose in the name and description. Tool-availability experiments may
restrict the available tools to measure that difference.

Keep CLI and MCP versions of shared user workflows equivalent when creating,
updating, or deleting scenarios. Requests and outcome assertions should agree;
transport-specific assertions need not. A genuinely entry-point-specific
experiment may have only one version if its purpose explains why. Do not
maintain a separate scenario inventory that can drift from the test files.

### Failure investigation

1. Check authentication, Excel availability, harness errors, and exhausted
   time/turn limits before assuming a product problem.
2. Inspect recorded calls and the actual workbook to locate the first failure.
3. Fix misleading guidance in the [canonical skill or description source](../skills/README.md#maintaining-skills-and-server-guidance), rebuild, and rerun.
4. For an implementation defect, add a deterministic regression test at the
   owning layer before fixing it.
5. Change the evaluation only when its request, setup, or assertion is wrong,
   not to conceal a missing feature or recovery path.

Do not mask product failures with skip/xfail. The shared harness may clearly
skip unavailable external prerequisites; a skip is not a passing evaluation.
Run only affected scenarios during iteration: Excel and external model access
are required, and model calls may incur costs.

### Independent saved-workbook checks

Chart-positioning and slicer scenarios share the same requests and checks for
both entry points. They reopen the saved file in a separate Excel instance with
macros/events disabled and links not updated, inspect it, and close without saving.
The checks compare real chart geometry, types, series, labels, slicer selections,
visible rows, and PivotTable totals. Expected answers are not embedded in prompts.

The reader and checks have deterministic tests, including deliberately wrong
saved positions and filters, that need Excel but no model or Azure access:

```powershell
python -m unittest discover -s . -p test_workbook_assertions.py -v
```

These tests also exercise the documented failure-aware batch and Power Query
recovery workflows. They use private CLI service pipes and temporary workbooks;
they do not stop another user's service.

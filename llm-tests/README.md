# ExcelMcp LLM Integration Tests

LLM-powered integration tests for both ExcelMcp MCP Server and Excel CLI using pytest-skill-engineering.

## Prerequisites

- Windows desktop with Microsoft Excel installed
- .NET 10 SDK
- Azure OpenAI endpoint configured
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

This installs the test dependencies from `pyproject.toml`, including `pytest-skill-engineering[copilot]`.

## Build MCP Server (Required)

```powershell
dotnet build ..\src\ExcelMcp.McpServer\ExcelMcp.McpServer.csproj -c Release
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

- `EXCEL_MCP_SERVER_COMMAND` — override MCP server command (full command line)
- `EXCEL_CLI_COMMAND` — override CLI command (default: `excelcli`)

Example:

```powershell
$env:EXCEL_MCP_SERVER_COMMAND = "d:\\source\\mcp-server-excel\\src\\ExcelMcp.McpServer\\bin\\Release\\net10.0-windows\\Sbroenne.ExcelMcp.McpServer.exe"
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
3. Fix misleading guidance in the [canonical skill or description source](../skills/README.md#maintaining-skills-and-mcp-prompts), rebuild, and rerun.
4. For an implementation defect, add a deterministic regression test at the
   owning layer before fixing it.
5. Change the evaluation only when its request, setup, or assertion is wrong,
   not to conceal a missing feature or recovery path.

Do not mask product failures with skip/xfail. The shared harness may clearly
skip unavailable external prerequisites; a skip is not a passing evaluation.
Run only affected scenarios during iteration: Excel and external model access
are required, and model calls may incur costs.

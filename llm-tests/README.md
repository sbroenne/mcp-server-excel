# ExcelMcp LLM Integration Tests

LLM-powered integration tests for both ExcelMcp MCP Server and Excel CLI using pytest-aitest.

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

## Setup (uv + local pytest-aitest)

From this directory:

```powershell
uv sync
```

This uses a local editable dependency via:

```toml
[tool.uv.sources]
pytest-aitest = { path = "../../../pytest-aitest", editable = true }
```

## Build MCP Server (Required)

```powershell
dotnet build ..\..\src\ExcelMcp.McpServer\ExcelMcp.McpServer.csproj -c Release
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
$env:EXCEL_MCP_SERVER_COMMAND = "d:\\source\\mcp-server-excel\\src\\ExcelMcp.McpServer\\bin\\Release\\net10.0\\ExcelMcp.McpServer.exe"
$env:EXCEL_CLI_COMMAND = "excelcli"
```

## Test Structure

- `test_mcp_*.py` — MCP Server workflows
- `test_cli_*.py` — CLI workflows
- `test_*calculation_mode*.py` — new calculation mode scenarios
- `Fixtures/` — shared test inputs (CSV/JSON/M files)
- `TestResults/` — HTML reports and artifacts

# Excel MCP Server - LLM Integration Tests

This project contains integration tests for the Excel MCP Server using [agent-benchmark](https://github.com/mykhaliev/agent-benchmark) - a framework for testing AI agents and MCP tool usage.

## Prerequisites

### 1. Build the Main Project First

The tests require the MCP server to be built. Run the following command from the repository root:

```powershell
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj -c Release
```

### 2. Azure OpenAI Configuration

These tests use Azure OpenAI for LLM interactions and MCP tool invocations. Configure the following environment variable:

- `AZURE_OPENAI_ENDPOINT` - Your Azure OpenAI endpoint URL

Authentication uses Entra ID (DefaultAzureCredential) - no API key needed.

### 3. Windows Desktop with Excel

These tests automate Excel via COM interop. They require:
- Windows desktop with UI access
- Microsoft Excel installed
- **NOT suitable for headless CI/CD pipelines**
- Run on a Windows machine with an active desktop session

### 4. Agent-Benchmark Tool

The PowerShell runner script will automatically download agent-benchmark on first run. Alternatively, you can:
- Download from [agent-benchmark releases](https://github.com/mykhaliev/agent-benchmark/releases)
- Build from source: `git clone https://github.com/mykhaliev/agent-benchmark && cd agent-benchmark && go build`
- Use a local Go project with `go run` mode (see Configuration below)

## Configuration

The test runner supports configuration via JSON files. Settings are loaded in this order (later overrides earlier):

1. `llm-tests.config.json` - Shared defaults (committed to repo)
2. `llm-tests.config.local.json` - Personal settings (git-ignored)
3. Command-line parameters - Override everything

### Configuration File

Create `llm-tests.config.local.json` for your personal settings:

```json
{
  "$schema": "./llm-tests.config.schema.json",
  "agentBenchmarkPath": "../../../../agent-benchmark",
  "agentBenchmarkMode": "go-run",
  "verbose": false,
  "build": false
}
```

### Configuration Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `agentBenchmarkPath` | string | `null` | Path to agent-benchmark (absolute or relative to test dir) |
| `agentBenchmarkMode` | string | `"executable"` | `"executable"` for .exe, `"go-run"` for Go project |
| `verbose` | boolean | `false` | Show detailed output |
| `build` | boolean | `false` | Build MCP server before tests |

### YAML Scenario Configuration

Test scenarios use minimal configuration with sensible defaults:

```yaml
providers:
  - name: azure-openai-gpt41
    type: AZURE
    auth_type: entra_id
    model: gpt-4.1
    baseUrl: "{{AZURE_OPENAI_ENDPOINT}}"
    version: 2025-01-01-preview
    retry:
      retry_on_429: true
      max_retries: 5

servers:
  - name: excel-mcp
    type: stdio
    command: "{{SERVER_COMMAND}}"
    server_delay: 30s  # Allow server initialization

agents:
  - name: gpt41-agent
    servers:
      - name: excel-mcp
    provider: azure-openai-gpt41
```

**Simplified Configuration:**
- No `rate_limits` needed - handled by agent-benchmark retry logic
- No custom `system_prompt` - defaults work well for tool calling
- No `clarification_detection` - simplified agent configuration
- Uses `server_delay: 30s` for reliable server initialization

## Running the Tests

### Using PowerShell Runner (Recommended)

```powershell
# Run all tests
.\Run-LLMTests.ps1 -Build

# Run a specific scenario
.\Run-LLMTests.ps1 -Scenario excel-file-worksheet-test.yaml
```

### Using agent-benchmark Directly

```powershell
# Build the server first
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj -c Release

# Run a scenario (after substituting {{SERVER_COMMAND}})
agent-benchmark `
    -f tests/ExcelMcp.McpServer.LLM.Tests/Scenarios/excel-file-worksheet-test.yaml `
    -o report `
    -reportType html,json `
    -verbose
```

## Project Structure

```
ExcelMcp.McpServer.LLM.Tests/
├── Scenarios/
│   ├── _config-template.yaml.template     # Reference configuration template
│   ├── excel-file-worksheet-test.yaml     # File and worksheet lifecycle
│   ├── excel-range-test.yaml              # Range data operations
│   ├── excel-table-test.yaml              # Table operations
│   ├── excel-slicer-test.yaml             # PivotTable and Table slicers
│   ├── excel-powerquery-datamodel-test.yaml  # Complete BI workflow
│   └── excel-modification-patterns-test.yaml # Incremental update validation
├── output/                                # Test artifacts (git-ignored)
├── Run-LLMTests.ps1                       # PowerShell test runner
├── llm-tests.config.json                  # Shared configuration defaults
├── llm-tests.config.local.json            # Personal settings (git-ignored)
├── llm-tests.config.schema.json           # JSON schema for config files
├── TestResults/                           # HTML reports (generated)
└── README.md                              # This file
```

## Test Scenarios

| Scenario | Tools Covered | Description |
|----------|---------------|-------------|
| `excel-file-worksheet-test.yaml` | excel_file, excel_worksheet | File lifecycle and worksheet operations |
| `excel-range-test.yaml` | excel_range | Get/set values, 2D array format |
| `excel-table-test.yaml` | excel_table | Table creation and data operations |
| `excel-slicer-test.yaml` | excel_slicer, excel_table, excel_pivottable | PivotTable and Table slicer operations |
| `excel-powerquery-datamodel-test.yaml` | excel_powerquery, excel_datamodel, excel_pivottable | Complete BI workflow |
| `excel-modification-patterns-test.yaml` | excel_range | Validates LLM uses incremental updates |

## Test Isolation

Each test uses unique temporary file paths to ensure isolation:
- Files use `{{randomValue type='UUID'}}` placeholders for unique names per run
- Tests create files in `C:/temp/` with descriptive names + UUID suffix
- No cleanup scenarios required - temp files are naturally isolated
- Each test session operates on its own file

## Writing Test Scenarios

### Test Consolidation for Token Optimization

Tests are consolidated into multi-step prompts to reduce token overhead:

**BEFORE (multiple tests):**
```yaml
- name: "Create Excel file"
  prompt: "Create a new Excel file..."
- name: "Open the file"
  prompt: "Open the file..."
- name: "List worksheets"
  prompt: "List worksheets..."
```

**AFTER (consolidated):**
```yaml
- name: "Complete file lifecycle"
  prompt: |
    1. Create a new empty Excel file at C:/temp/test-{{randomValue type='UUID'}}.xlsx
    2. Open the file
    3. List all active Excel sessions
    4. Close the file without saving
```

**Benefits:**
- Reduced token consumption per test run
- Fewer round-trips to the LLM
- More realistic multi-step workflows
- File path UUIDs ensure test isolation

### User Prompts: Use Natural Language

**BAD - Leading the witness:**
```yaml
prompt: "Use excel_file with action 'create-empty' and excelPath 'C:/temp/test.xlsx'"
```

**GOOD - Natural user request:**
```yaml
prompt: "Create a new empty Excel file at C:/temp/test.xlsx"
```

### Assertion Types

| Assertion | Description |
|-----------|-------------|
| `tool_called` | Verifies specific MCP tool was invoked |
| `tool_param_equals` | Verifies tool parameter values |
| `output_regex` | Regex match on LLM response |
| `no_hallucinated_tools` | Ensures only real tools called |
| `max_latency_ms` | Performance threshold |
| `anyOf` | OR logic for multiple valid approaches |
| `allOf` | AND logic for required conditions |

## Test Reports

HTML reports are generated in `TestResults/` directory after each test run. Reports include:
- Scenario results (pass/fail)
- Step-by-step execution details
- Tool calls and responses
- Assertion results

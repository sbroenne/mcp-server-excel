# Excel CLI - LLM Integration Tests

This directory contains LLM-based integration tests for the Excel CLI tool (`excelcli`). These tests use [agent-benchmark](https://github.com/mykhaliev/agent-benchmark) to verify that AI agents can effectively use the CLI to automate Excel tasks.

## Overview

Unlike MCP Server tests (which test the Model Context Protocol integration), these tests verify:
- AI agents can discover and use CLI commands via `--help`
- Proper use of quiet mode (`-q`) for JSON output
- Session management workflow (open → operations → close)
- Error handling and recovery via CLI

## Prerequisites

- Windows 10/11 with Microsoft Excel installed
- .NET 10 SDK
- Azure OpenAI endpoint (set `AZURE_OPENAI_ENDPOINT` environment variable)
- agent-benchmark (auto-downloaded or use local build)

## Running Tests

From the repository root:

```powershell
# Run all CLI tests
.\scripts\Run-LLMTests.ps1 -Component cli

# Run specific scenario
.\scripts\Run-LLMTests.ps1 -Component cli -Scenario excel-range-cli-test.yaml

# Build CLI before running tests
.\scripts\Run-LLMTests.ps1 -Component cli -Build

# Use local agent-benchmark build
.\scripts\Run-LLMTests.ps1 -Component cli -AgentBenchmarkPath "D:\source\agent-benchmark"
```

## Configuration

Create `llm-tests.config.local.json` (git-ignored) for personal settings:

```json
{
  "agentBenchmarkPath": "D:/source/agent-benchmark",
  "agentBenchmarkMode": "go-run",
  "verbose": false,
  "build": true
}
```

## Test Scenarios

| Scenario | Description |
|----------|-------------|
| `excel-chart-positioning-cli-test.yaml` | Chart positioning without data overlap |
| `excel-file-worksheet-cli-test.yaml` | File and worksheet management |
| `excel-modification-patterns-cli-test.yaml` | Targeted updates vs delete-rebuild |
| `excel-pivottable-layout-cli-test.yaml` | PivotTable layout styles |
| `excel-powerquery-datamodel-cli-test.yaml` | Power Query and Data Model workflow |
| `excel-range-cli-test.yaml` | Range operations (get/set values) |
| `excel-slicer-cli-test.yaml` | PivotTable and Table slicers |
| `excel-table-cli-test.yaml` | Table operations (create, query) |

## Agent Skills

Tests use the `excel-cli` skill from `skills/excel-cli/` which provides:
- CLI command syntax and examples
- Session workflow patterns
- Best practices for quiet mode JSON output

## Directory Structure

```
ExcelMcp.CLI.LLM.Tests/
├── llm-tests.config.json     # Shared configuration
├── Scenarios/                # Test scenario files
│   ├── excel-range-cli-test.yaml
│   ├── excel-table-cli-test.yaml
│   └── ...
├── TestResults/              # Test output (git-ignored)
└── Fixtures/                 # Test data files

scripts/
└── Run-LLMTests.ps1          # Central test runner (use -Component cli)
```

## Related

- [MCP Server LLM Tests](../ExcelMcp.McpServer.LLM.Tests/) - Tests for MCP Server
- [Excel CLI Skill](../../skills/excel-cli/) - Agent skill for CLI
- [agent-benchmark](https://github.com/mykhaliev/agent-benchmark) - Test framework

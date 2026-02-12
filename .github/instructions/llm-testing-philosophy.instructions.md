---
applyTo: "llm-tests/**"
---

# LLM Testing Philosophy

> **⚠️ CORE PRINCIPLE: Tests simulate real users. Failures expose product gaps, not test gaps.**

## What Are LLM Tests?

LLM tests use an AI coding agent (with its own default system prompt) to exercise our CLI and MCP tools. The agent receives a **natural language prompt** — the same kind a real user would type — and must figure out how to accomplish the task using our tools.

These tests do NOT test the LLM. They test **our product's usability surface**:
- Skill documentation (SKILL.md)
- CLI `--help` output
- MCP tool descriptions (XML `/// <summary>`)
- Error messages and recovery hints
- Parameter naming and discoverability
- Workflow coherence

## The Golden Rule

> **If the LLM can't figure it out, fix the product — never fix the test.**

When an LLM test fails, the root cause is ALWAYS one of:
1. Our **skill docs** don't explain the workflow clearly enough
2. Our **tool descriptions** are misleading or incomplete
3. Our **CLI --help** output doesn't show the right examples
4. Our **error messages** don't guide recovery
5. Our **parameter names** are confusing or undiscoverable
6. The test itself is unreasonable (rare — only fix if a human couldn't do it either)

## What NEVER Belongs in a Test

### ❌ CLI Command Guidance in Prompts

A real user doesn't know our CLI syntax. Neither should the test prompt.

```python
# ❌ WRONG: Teaching the LLM how to use our CLI
prompt = """
Create a PivotTable then set layout to Compact using 
'excelcli pivottablecalc' (run --help to see options).
The compact layout uses row-layout value 0.
"""

# ✅ CORRECT: Natural user request
prompt = """
Create a PivotTable with Compact layout showing
Department and Team as rows, Hours as values.
"""
```

### ❌ Tool-Specific Hints in Prompts

```python
# ❌ WRONG: Directing the LLM to a specific command
prompt = """
Use the 'excelcli chartconfig' command to change the chart title.
"""

# ✅ CORRECT: What the user wants
prompt = """
Change the chart title to "Q1 Sales Report" without 
deleting and recreating the chart.
"""
```

### ❌ System Prompt Engineering

The system prompt belongs to the agent, not to us. We don't control what system prompt VS Code, Cursor, Claude Desktop, or any other host uses.

```python
# ❌ WRONG: Adding our own guidance to the system prompt
agent = Agent(
    system_prompt=(
        "Run 'excelcli <command> --help' when unsure about parameter names\n"
        "Always use -q flag for clean JSON output"
    ),
)

# ✅ CORRECT: No system prompt (use skill only) or minimal role context
agent = Agent(
    skill=excel_cli_skill,  # Our product — this IS the right place for guidance
)
```

### ❌ Session Recovery Instructions

```python
# ❌ WRONG: Teaching the LLM our session model
prompt = """
IMPORTANT: First, run 'excelcli session list' to confirm 
which file is open, then use that file path.
"""

# ✅ CORRECT: If session discovery is hard, fix the product:
# - Better error messages: "No active session. Run 'excelcli session list' to see open files."
# - Better skill docs: Add "Session Management" section with recovery patterns
# - Better --help: Show session workflow in command help
```

## What DOES Belong in a Test

### Natural Language Prompts

Write prompts as a knowledgeable Excel user would. They know Excel concepts but NOT our specific CLI/MCP tool syntax.

```python
prompt = f"""
Create a new Excel file at {unique_path('sales-analysis')}

Enter this sales data:
Region, Product, Sales
North, Widget, 15000
South, Gadget, 12000

Create a PivotTable showing Region as rows and Sum of Sales as values.
Add a slicer for the Region field.
Filter to show only North region.

Save and close the file.
"""
```

### Reasonable Assertions

Assert on **outcomes** the user cares about, not implementation details:

```python
# ✅ Good: Did the operation succeed?
assert result.success
assert_cli_exit_codes(result)

# ✅ Good: Did the LLM report key results?
assert_regex(result.final_response, r"(?i)(pivot|region|north)")

# ⚠️ Fragile: Exact numeric values across 5 conversation turns
assert_regex(result.final_response, r"\$?43,500\.00")  # Requires perfect execution of ALL prior steps
```

### Appropriate Complexity

Match test complexity to what a real user would attempt in one conversation:

| Complexity | Turns | Example |
|-----------|-------|---------|
| Simple | 1 | Create file → write data → read back → close |
| Medium | 1-2 | Create file → build table → add chart → save |
| Complex | 2-3 | Build data model → create measures → analyze |
| Unreasonable | 5+ | 13-step workflow with exact numeric assertions on final state |

## Where to Fix Failures

When a test fails, investigate in this order:

### 1. Skill Documentation (`skills/excel-cli/SKILL.md`, `skills/excel-mcp/SKILL.md`)

The skill IS our product's interface to LLMs. If the LLM doesn't know how to:
- Set a PivotTable layout → Add workflow patterns to the skill
- Use `--values-file` instead of `--values` → Document when to use file params
- Discover `chartconfig` commands → Add chart modification patterns

```markdown
## Chart Modification Patterns
To change chart properties WITHOUT deleting the chart:
- Title: `chartconfig set-title`
- Type: `chartconfig set-type`
```

### 2. CLI `--help` Output

The `--help` text is what the LLM sees when it runs `excelcli <command> --help`. If a parameter is hard to discover, improve the help text.

Look at:
- `[Description]` attributes on Settings properties
- Parameter ordering and grouping
- Whether common workflows are clear from help alone

### 3. MCP Tool Descriptions (`/// <summary>` in Tool classes)

For MCP tests, the tool description IS the API documentation. If the LLM picks wrong tools or wrong parameters, improve the XML docs.

### 4. Error Messages

When a CLI command fails, does the error message tell the LLM how to fix it? 

```
# ❌ Bad error message
"Parameter 'mCode' is required"

# ✅ Good error message  
"Parameter 'mCode' is required for 'create' action. Provide --m-code with inline M code or --m-code-file with a file path."
```

### 5. Parameter Naming

If the LLM consistently gets a parameter name wrong, the name might be confusing. Consider renaming or adding aliases.

## Test Structure Conventions

### Agent Configuration

```python
agent = Agent(
    name="descriptive-test-name",
    provider=Provider(model=f"azure/{DEFAULT_MODEL}", rpm=DEFAULT_RPM, tpm=DEFAULT_TPM),
    cli_servers=[excel_cli_server],   # CLI tests
    # OR
    mcp_servers=[excel_mcp_server],   # MCP tests
    skill=excel_cli_skill,            # Our product documentation
    max_turns=DEFAULT_MAX_TURNS,      # Always set explicitly
)
```

- **Always set `max_turns`** explicitly — relying on defaults creates silent failures
- **No custom `system_prompt`** unless testing a specific user persona (rare)
- **Always include `skill`** — this is our product's documentation

### Multi-Turn Tests

Each turn should be a natural continuation of the conversation:

```python
# Turn 1: Set up
result = await aitest_run(agent, "Create file and enter data...")
messages = result.messages

# Turn 2: Analyze (natural continuation)
result = await aitest_run(agent, "Now create a PivotTable from that data...", messages=messages)
```

Keep multi-turn tests to **2-3 turns maximum**. If you need 5 turns, the test is testing too many features at once — split it into separate tests.

## MCP/CLI Test Sync Rule (CRITICAL)

**Every test scenario MUST exist for BOTH CLI and MCP.** The two entry points are equal citizens — if a scenario is worth testing for one, it's worth testing for the other.

### Current Test Mapping

| Test Scenario | CLI File | MCP File |
|---------------|----------|----------|
| calculation_mode | `test_cli_calculation_mode.py` | `test_mcp_calculation_mode.py` |
| chart | `test_cli_chart.py` | `test_mcp_chart.py` |
| chart_positioning | `test_cli_chart_positioning.py` | `test_mcp_chart_positioning.py` |
| file_worksheet | `test_cli_file_worksheet.py` | `test_mcp_file_worksheet.py` |
| financial_report_automation | `test_cli_financial_report_automation.py` | `test_mcp_financial_report_automation.py` |
| modification_patterns | `test_cli_modification_patterns.py` | `test_mcp_modification_patterns.py` |
| pivottable_layout | `test_cli_pivottable_layout.py` | `test_mcp_pivottable_layout.py` |
| powerquery_datamodel | `test_cli_powerquery_datamodel.py` | `test_mcp_powerquery_datamodel.py` |
| range | `test_cli_range.py` | `test_mcp_range.py` |
| sales_report_workflow | `test_cli_sales_report_workflow.py` | `test_mcp_sales_report_workflow.py` |
| slicer | `test_cli_slicer.py` | `test_mcp_slicer.py` |
| table | `test_cli_table.py` | `test_mcp_table.py` |

### Rules for Creating / Updating / Deleting Tests

1. **Creating a new test:** ALWAYS create BOTH the CLI and MCP version. Name them identically except for the prefix (`test_cli_` vs `test_mcp_`). The prompt text should be identical — only the agent configuration differs (cli_servers vs mcp_servers, cli skill vs mcp skill).

2. **Updating a test:** Update BOTH the CLI and MCP version. If you change the prompt or assertions in one, apply the same change to the other. The test scenario must remain equivalent.

3. **Deleting a test:** Delete BOTH versions. Never leave an orphaned test in only one entry point.

4. **Test parity check:** Before committing, verify that `llm-tests/cli/` and `llm-tests/mcp_tests/` have matching test files (same scenarios, same count minus the prefix).

### Agent Configuration Differences

The ONLY differences between CLI and MCP versions of a test should be:

```python
# CLI version
agent = Agent(
    name="test-name-cli",
    cli_servers=[excel_cli_server],
    skill=excel_cli_skill,
    ...
)

# MCP version
agent = Agent(
    name="test-name-mcp",
    mcp_servers=[excel_mcp_server],
    skill=excel_mcp_skill,
    ...
)
```

Prompts, assertions, and test logic should be **identical**.

## SKILL.md Generation Awareness (CRITICAL)

**SKILL.md files are auto-generated. Never edit them directly.**

### Generation Pipeline

```
C# Interfaces (XML /// docs)
  → Roslyn Source Generator (ServiceRegistryGenerator)
    → _SkillManifest.g.cs (JSON const)
      → MSBuild Task (GenerateSkillFile.cs)
        → Scriban Templates (.sbn)
          → Generated SKILL.md
```

### Where to Fix What

| Problem | Fix Location | NOT Here |
|---------|-------------|----------|
| Wrong tool/command description | `I*Commands.cs` XML `/// <summary>` | `SKILL.md` |
| Wrong parameter docs | `I*Commands.cs` XML `/// <param>` | `SKILL.md` |
| Wrong skill prose/rules/workflows | `skills/templates/SKILL.cli.sbn` or `SKILL.mcp.sbn` | `SKILL.md` |
| Wrong reference doc content | `skills/shared/*.md` | `skills/excel-*/references/*.md` |
| Wrong Tool Selection table (MCP) | `skills/templates/SKILL.mcp.sbn` | `SKILL.md` |

### Key Files

- **Templates:** `skills/templates/SKILL.cli.sbn`, `skills/templates/SKILL.mcp.sbn`
- **Reference docs (source of truth):** `skills/shared/*.md` → copied to `skills/excel-*/references/`
- **Generated files (NEVER edit):** `skills/excel-cli/SKILL.md`, `skills/excel-mcp/SKILL.md`
- **Build command:** `dotnet build -c Release` regenerates SKILL.md and copies references

### Testing Impact

When a test fails because the LLM misuses a tool or parameter:
1. Check the SKILL template (`skills/templates/SKILL.*.sbn`) — is the guidance correct?
2. Check the reference doc (`skills/shared/*.md`) — are the parameter values correct?
3. Check the C# interface XML docs — is the tool description accurate?
4. Fix in the source file, rebuild (`dotnet build -c Release`), then re-run tests.

**Never fix a SKILL.md directly** — the fix will be lost on next build.

## Quick Checklist

Before submitting an LLM test PR:

- [ ] Prompts read like natural user requests, not CLI tutorials
- [ ] No CLI command names, flag names, or syntax in prompts
- [ ] No system prompt (or minimal role-setting only)
- [ ] `max_turns` set explicitly on Agent
- [ ] Assertions check outcomes, not implementation details
- [ ] At most 2-3 conversation turns
- [ ] If a test needs "hints" to pass, the hint belongs in the skill/help/tool docs instead
- [ ] **BOTH CLI and MCP versions of the test exist** (sync rule)
- [ ] **No direct edits to generated SKILL.md files** (template awareness)

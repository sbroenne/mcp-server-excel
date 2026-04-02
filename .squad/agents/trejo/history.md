# Trejo — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** Docs Lead
- **Joined:** 2026-03-15T10:42:22.625Z

## Learnings

### Calculation Mode LLM Discoverability Alignment (2026-03-XX)

**Task:** Fix docs/skills to align with repaired LLM tests validating calculation_mode tool discovery (both with and without skill guidance)

**Tests Expect:**
- LLM should call `calculation_mode` when writing 10+ cells (batch operations)
- Tool should be discoverable from tool description alone (without skill guidance)
- Workflow: set manual → write data → calculate → set automatic

**Root Cause of Drift:**
- Tool description was generic ("Set or get Excel calculation mode")
- No reference file dedicated to calculation mode (only inline in SKILL.md main body)
- LLMs lacked clear threshold guidance (10+ cell rule) in discoverable locations

**Fixes Applied:**
1. **Created `skills/excel-mcp/references/calculation.md`** — Dedicated reference explaining:
   - When to use (10+ cells threshold explicitly stated)
   - When NOT needed (small edits, reading formulas, immediate results)
   - 4-step workflow with clear purpose of each step
   - Scenario examples (sales table, dashboard) showing performance gains
2. **Enhanced tool description in CalculationModeCommands.cs** — Now includes:
   - "Optimize bulk write performance" → immediate relevance signaling
   - "10+ cells" threshold explicitly mentioned
   - "BATCH WORKFLOW (required for 10+ cell operations)" emphasizes when to use
   - "NOT needed for: reading formulas, small edits (1-9 cells)" eliminates false positives
3. **Updated SKILL.mcp.sbn template** — Added calculation.md to reference documentation list (alphabetically after anti-patterns)

**Outcome:**
- Tool description auto-generates to MCP tool signatures at build time (generators read Core layer attributes)
- SKILL.md regenerates from template, auto-includes calculation.md link
- Skills copied to references/ folder at build time (both MCP clients and skill-based clients get identical guidance)
- Single source of truth: Core attributes → generated MCP tools → SKILL.md → skill references

**Pattern Observed:**
- MCP tool descriptions drive LLM discovery when skill docs unavailable
- Concrete thresholds (10+, 1-9) are more discoverable than vague language ("many")
- Explicit "NOT needed for" sections prevent false positives
- Dedicated reference files improve hierarchical organization (especially for complex tools)

### Claude Desktop Release Artifacts & Documentation (2026-03-15)

**What we release for Claude Desktop:**
- `.mcpb` bundle file (e.g., `excel-mcp-1.7.0.mcpb`) built via `mcpb/Build-McpBundle.ps1`
- Self-contained Windows x64 executable (`excel-mcp-server.exe`) embedded in the bundle
- The bundle contains: manifest.json (with metadata), icon-512.png, README.md, LICENSE, CHANGELOG.md, and /server/ folder with the exe
- Built and uploaded by `release.yml` Job 4 (`build-mcpb`) to GitHub release artifacts
- Installation: users download from latest release and double-click to install in Claude Desktop

**What docs say about Claude Desktop:**
- Main README: Quick start table lists "Claude Desktop" with link to download `.mcpb` from releases
- `mcpb/README.md`: End-user facing documentation (ships in the bundle) — comprehensive, 120+ lines with examples, requirements, troubleshooting
- `skills/excel-mcp/references/claude-desktop.md`: Configuration guide for manual setup (Windows container considerations, file system access, session persistence)
- `docs/CLAUDE-MCPB-SUBMISSION.md`: Developer guide for Anthropic submission process (checklist, manifest schema, asset requirements)
- Release workflow notes say MCPB is released alongside MCP Server, CLI, VS Code Extension, and Agent Skills
- CHANGELOG.md lists MCPB as one of four components: "- **MCPB** - Claude Desktop bundle for one-click installation"

**Manifest metadata:**
- `mcpb/manifest.json` (v0.3 spec): display_name="Excel (Windows)", 23 tools, 214 operations listed, Claude Desktop >=0.10.0, platforms: win32 only
- Note: Manifest says 23 tools, 214 ops; README says 25 tools, 230 ops — count mismatch

**Release mechanics check:**
- Release workflow downloads artifacts from 5 build jobs, generates release notes with installation instructions
- Release notes explicitly mention "Claude Desktop (MCPB)" as one of four installation options, with download link
- Docs match actual release: `.mcpb` file is what gets shipped

**Key gap identified:**
- Manifest.json says 23 tools/214 operations, but code and other docs consistently say 25 tools/230 operations
- This is a docs consistency issue, not a release mechanics issue

### LLM Tests Migration Write-Up (2026-03-XX)

**Task:** Draft polished GitHub issue and PR copy for pytest-aitest → pytest-skill-engineering migration

**Deliverables Created:**
1. **GitHub Issue** — Title + body covering scope, changes, validation, impact
2. **GitHub PR** — Title + body with technical details, checklist, validation steps
3. **CHANGELOG Entry** — "Added" section with philosophy emphasis and scope summary

**Key Narrative Choices:**
- Lead with **framework scope** not patches: "Clean full rewrite, no backward-compat shim"
- Emphasize **philosophy shift**: "Tests validate skill docs quality, not LLM capability"
- Highlight **Golden Rule**: "Fix product issues, never hide test failures"
- Assert **no regression**: "All 12 test scenarios preserved with enhanced assertions"

**Documentation Insights:**
- LLM testing philosophy is canonical source for test patterns (instruction docs)
- Natural language prompts (user perspective, not CLI tutoring) is non-negotiable
- CLI/MCP test parity enforced by "MCP/CLI Sync Rule" (same scenarios for both)
- Failed tests = product issues (skill docs, tool descriptions, error messages), never test brittleness

**Migration Pattern:**
- Dependencies: pytest-aitest → pytest-skill-engineering[copilot] >= 0.5.9
- Fixtures: CopilotEval-based (conftest.py rewrite)
- Test setup: GitHub Copilot auth gating, explicit max_turns=20, 600s timeout
- Manual execution: `uv run pytest -m mcp|cli|aitest` (not CI/CD automated)

**Governance Implication:**
- Skill docs are auto-synced to MCP prompts at build time (single source of truth)
- Test failures always trigger skill/doc/help improvements, not test hides (xfail/skip forbidden)
- Both CLI and MCP are equal citizens in test validation

- 2026-04-06: **Calculation Mode LLM Discoverability Alignment Complete.** Created dedicated `skills/excel-mcp/references/calculation.md` reference explaining when/when-not to use (10+ cells threshold explicit, scenario examples, best practices). Enhanced `CalculationModeCommands.cs` tool description with "Optimize bulk write performance," explicit 10+ cell threshold, "BATCH WORKFLOW (required)" emphasizing workflow requirement. Updated `skills/templates/SKILL.mcp.sbn` to include calculation.md in reference list. Single source of truth flow: Core → MCP schema → Claude Desktop tool description; SKILL template → skill doc → reference docs. Test alignment: with-skill tests read calculation.md, no-skill tests read enhanced tool description. Both scenarios now equally discoverable. No product code changes — pure skill/doc alignment.


- 2026-04-02: **Pre-commit release-gate alignment.** The VS Code extension mismatch (`@types/vscode` 1.110 vs `engines.vscode` 1.109) only surfaced at `npm run package`; `npm install` and TypeScript compile were not enough. When hardening `scripts\pre-commit.ps1`, keep the hook inventory in docs synchronized with the actual script and use the release packaging path itself when a release blocker lives in manifest/package metadata.

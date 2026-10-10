# Squad Team

> sbroenne-silver-robot

## Coordinator

| Name | Role | Notes |
|------|------|-------|
| Squad | Coordinator | Routes work, enforces handoffs and reviewer gates. |

## Members

| Name | Role | Charter | Status |
|------|------|---------|--------|
| Lead | Lead | `.squad/agents/lead/charter.md` | ✅ Active |
| Runtime Engineer | Core and Excel COM Engineer | `.squad/agents/runtime-engineer/charter.md` | ✅ Active |
| Entry Points Engineer | Service, CLI, and MCP Engineer | `.squad/agents/entry-points-engineer/charter.md` | ✅ Active |
| Quality Engineer | Tests and Regression Evidence | `.squad/agents/quality-engineer/charter.md` | ✅ Active |
| Docs & Extension Engineer | Documentation, Guidance, and Extension | `.squad/agents/docs-extension-engineer/charter.md` | ✅ Active |
| Scribe | Decision Merger | `.squad/agents/scribe/charter.md` | 📋 Silent |
| Ralph | Work Monitor | `.squad/agents/ralph/charter.md` | 🔄 Monitor |
| Rai | RAI Reviewer | `.squad/agents/Rai/charter.md` | 🛡️ On-demand |
| Fact Checker | Verifier and Devil's Advocate | `.squad/agents/fact-checker/charter.md` | 🔍 On-demand |


## Coding Agent

<!-- copilot-auto-assign: false -->

| Name | Role | Charter | Status |
|------|------|---------|--------|
| @copilot | Coding Agent | — | 🤖 Coding Agent |

### Capabilities

**🟢 Good fit — auto-route when enabled:**
- Bug fixes with clear reproduction steps
- Test coverage (adding missing tests, fixing flaky tests)
- Lint/format fixes and code style cleanup
- Dependency updates and version bumps
- Small isolated features with clear specs
- Boilerplate/scaffolding generation
- Documentation fixes and README updates

**🟡 Needs review — route to @copilot but flag for squad member PR review:**
- Medium features with clear specs and acceptance criteria
- Refactoring with existing test coverage
- API endpoint additions following established patterns
- Migration scripts with well-defined schemas

**🔴 Not suitable — route to squad member instead:**
- Architecture decisions and system design
- Multi-system integration requiring coordination
- Ambiguous requirements needing clarification
- Security-critical changes (auth, encryption, access control)
- Performance-critical paths requiring benchmarking
- Changes requiring cross-team discussion

## Project Context

- **Owner:** sbroenne
- **Project:** ExcelMcp (sbroenne/mcp-server-excel) — Windows-only automation of installed desktop Excel through COM; the MCP Server and `excelcli` are equal entry points over shared Core commands
- **Stack:** C#/.NET (SDK from `global.json`), Excel COM, MCP SDK, generated Service/CLI/MCP surfaces, PowerShell scripts, MkDocs website, VS Code extension (TypeScript)
- **Rules:** [AGENTS.md](../AGENTS.md) and the guides it maps; system map in [CONTEXT.md](../CONTEXT.md)
- **Project directory name:** sbroenne-silver-robot
- **Created:** 2026-10-06

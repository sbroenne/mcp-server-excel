# Work Routing

How to decide who handles what.

## Routing Table

| Work Type | Route To | Examples |
|-----------|----------|----------|
| COM interop, Excel automation, Core commands | Shiherlis | Add/fix Core command, COM cleanup, batch.Execute patterns, ExcelWriteGuard |
| **COM API review (MANDATORY)** | **Hanna** | **Every change touching COM/Excel — validates against docs, never guesses** |
| MCP Server tools, CLI commands, service layer | Cheritto | Add MCP tool action, CLI parity, source generators, ForwardToService |
| Integration tests, test fixtures, quality | Nate | Write tests, fix flaky tests, round-trip validation, TDD red-green |
| Skills, docs, READMEs, CHANGELOG, release | Trejo | Update SKILL.md, README counts, FEATURES.md, gh-pages, CHANGELOG |
| Copilot CLI plugin packaging, `*.agent.md`, `hooks.json`, plugin `.mcp.json`, marketplace submission | Kelso | Bundle skills+MCP as Copilot CLI plugin, submit to github/copilot-plugins |
| Architecture, API design, exception patterns | McCauley | Review PRs, architectural decisions, pattern enforcement |
| Code review | McCauley | Review PRs, check quality, enforce Critical Rules |
| Testing | Nate | Write tests, find edge cases, verify fixes |
| Scope & priorities | McCauley | What to build next, trade-offs, decisions |
| Session logging | Scribe | Automatic — never needs routing |

## Issue Routing

| Label | Action | Who |
|-------|--------|-----|
| `squad` | Triage: analyze issue, assign appropriate `squad:{member}` label, and comment with triage notes | Lead |
| `squad:{name}` | Pick up issue and complete the work | Named member |

### How Issue Assignment Works

1. When a GitHub issue gets the `squad` label, the **Lead** triages it — analyzing content, matching it to the appropriate squad member's domain, assigning the corresponding `squad:{member}` label, and commenting with triage notes.
2. When a `squad:{member}` label is applied, that member picks up the issue in their next session.
3. Members can reassign by removing their label and adding another member's label.
4. The `squad` label is the "inbox" — untriaged issues waiting for Lead review.

## Rules

1. **Eager by default** — spawn all agents who could usefully start work, including anticipatory downstream work.
2. **Scribe always runs** after substantial work, always as `mode: "background"`. Never blocks.
3. **Quick facts → coordinator answers directly.** Don't spawn an agent for "what port does the server run on?"
4. **When two agents could handle it**, pick the one whose domain is the primary concern.
5. **"Team, ..." → fan-out.** Spawn all relevant agents in parallel as `mode: "background"`.
6. **Anticipate downstream work.** If a feature is being built, spawn the tester to write test cases from requirements simultaneously.
7. **Issue-labeled work** — when a `squad:{member}` label is applied to an issue, route to that member. The Lead handles all `squad` (base label) triage.

## Domain → Agent (Quick Reference)

| Domain | Primary | Secondary |
|--------|---------|----------|
| Architecture, code review, decisions | McCauley | — |
| COM interop, Excel automation | Shiherlis | Hanna (mandatory review) |
| COM API correctness, documentation validation | Hanna | — |
| MCP Server, CLI parity | Cheritto | Hanna (if COM changes) |
| Integration tests, quality | Nate | — |
| Skills, docs, release | Trejo | — |
| Copilot CLI plugin packaging & distribution | Kelso | Trejo (skill content), Cheritto (MCP wiring) |

## Mandatory Review Gate

**Hanna (COM Interop Expert) MUST review ANY change that:**
- Adds, modifies, or removes COM interop calls (`dynamic` Excel objects)
- Changes COM object lifecycle patterns (acquisition, release, cleanup)
- Introduces new Excel Object Model usage (new properties, methods, collections)
- Modifies `batch.Execute()` lambdas in Core commands
- Touches `src/ExcelMcp.ComInterop/` infrastructure
- Changes Excel session management or shutdown patterns

**Workflow:** Shiherlis writes COM code → Hanna validates against documentation → McCauley approves architecture.


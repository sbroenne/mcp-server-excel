# Work Routing

How to decide who handles what.

## Routing Table

| Work Type | Route To | Examples |
|-----------|----------|----------|
| Scope, architecture, cross-area integration, code review, unmatched work | Lead | Design disagreements, ADR conflicts, review of a finished change |
| Core commands, ComInterop, Excel COM sessions and cleanup | Runtime Engineer | COM leaks, timeouts, session recovery, Power Query/PivotTable/chart behavior |
| Service, CLI, MCP Server, generators, entry-point contracts | Entry Points Engineer | Tool schemas, CLI flags, generated routing, MCP and CLI disagreeing |
| .NET tests, regression evidence, Excel E2E, check scripts | Quality Engineer | Failing regression test, sequential `scripts\Test-E2E.ps1` run, flaky test |
| Docs, product skills and agent guidance, website, VS Code extension, changesets | Docs & Extension Engineer | Stale docs, skill sources, `gh-pages`, `vscode-extension/` |

### Built-in support (not owners of product work)

| Member | Use when |
|--------|----------|
| Scribe | Merging decision inbox files into `decisions.md` after substantial work |
| Ralph | Checking the work queue or backlog |
| Rai | Responsible AI review, only when requested or selected as a gate |
| Fact Checker | Verifying claims, or challenging a plan or conclusion before it is accepted |

### @copilot

Opt-in only: `copilot-auto-assign` is `false` in `.squad/team.md`. Route to @copilot only when asked, using the capability profile there.

## Issue Routing

| Label | Action | Who |
|-------|--------|-----|
| `squad` | Triage: analyze issue, assign `squad:{member}` label | Lead |
| `squad:{name}` | Pick up issue and complete the work | Named member |

### How Issue Assignment Works

1. When a GitHub issue gets the `squad` label, the **Lead** triages it — analyzing content, assigning the right `squad:{member}` label, and commenting with triage notes.
2. When a `squad:{member}` label is applied, that member picks up the issue in their next session.
3. Members can reassign by removing their label and adding another member's label.
4. The `squad` label is the "inbox" — untriaged issues waiting for Lead review.

## Rules

1. **One accountable owner per task.** Pick the member whose area is the primary concern; the Lead owns anything unmatched.
2. **Bring in others only for a concrete need** (a changed contract, a review, test evidence), and keep the collaboration bounded.
3. **Changes to Core contracts, generators, or entry-point behavior** involve the Entry Points Engineer so MCP Server and `excelcli` stay in agreement.
4. **Excel-dependent tests run sequentially**; never run two in parallel.
5. **Scribe runs after substantial work**, in the background, and never blocks.
6. **Quick facts → coordinator answers directly.** Don't spawn an agent for a simple lookup.
7. **Issue-labeled work** — a `squad:{member}` label routes to that member; the Lead handles the base `squad` label.
8. `llm-tests/` evaluations run only when the user explicitly asks.
# Session Log: 20260423T154955Z - Copilot CLI Plugin Planning

**Requester:** Stefan Brönner  
**Scope:** Define shape and execution plan for ExcelMcp Copilot CLI plugins  
**Outcome:** Two plugins finalized (excel-mcp + excel-cli), all 6 architectural decisions locked, GitHub Action automation approved

---

## Planning Arc

### Phase 1: User Directive Capture (Stefan)

**20260423T164049Z**  
Stefan clarified plugin scope: **GitHub Copilot CLI plugins** (not agentskills.io, not MCPB bundles, not VS Code extensions).
- Entry point: https://docs.github.com/en/copilot/concepts/agents/copilot-cli/about-cli-plugins
- Contents: agents (*.agent.md), skills (SKILL.md), hooks (hooks.json), MCP configs (.mcp.json), LSP configs
- Distribution: Copilot CLI plugin marketplaces (e.g., github/copilot-plugins, github/awesome-copilot)

### Phase 2: Team Hiring & Initial Plan (Copilot → Kelso)

Copilot spawned Kelso as Copilot CLI Plugin Engineer to draft initial plugin plan. 

Kelso produced `initial-plugin-plan.md` with 7 decisions:
1. Custom Excel agent? (TBD)
2. MCP binary distribution? (TBD)
3. Marketplace submission? (TBD)
4. Plugin layout? (TBD)
5. Version pinning? (TBD)
6. Publication mechanism? (TBD)
7. Windows-only gating? (TBD)

### Phase 3: User Decisions (Stefan) — 20260423T172300Z

Stefan locked 5 decisions:

| # | Decision | Choice |
|---|----------|--------|
| 6 | Plugin layout | TWO separate plugins (excel-mcp for MCP, excel-cli for CLI) |
| 4 | MCP binary distribution | Bundle with plugin (not separate dotnet tool) |
| 5 | Marketplace submission | Defer to v2 (not submitting to github/copilot-plugins in v1) |
| 7 | Publication mechanism | Automated via GitHub Action (deviates from office-coding-agent manual precedent) |

**Still open** (left to Kelso):
- #1 Plugin names
- #2 Published repo name
- #3 Custom Excel agent (yes/no + rationale)

### Phase 4: Kelso Refinement — 20260423T154955Z

Kelso refined the plan and **pushed back on decision #5** (MCP binary distribution):

**Kelso's Override:**
> Instead of bundling the MCP binary IN Git, ship a `bin/download.ps1` script that pulls the binary from the matching GitHub Release tag at user install time.

**Rationale:**
- .NET self-contained publish = 50–80MB typical for Windows x64
- Git repos struggle with large binaries (slow clones, bloat, unrecoverable history)
- Each release = +50–80MB to Git history (chronic bloat)
- Release download = lean repo, fast clones, clean maintenance

**Tradeoff:** Two-step install (plugin + binary download) vs. single-step with chronic Git bloat.
→ **Kelso won the argument: Release download is the better long-term strategy.**

Kelso also delivered:
- ✅ Answered #1: Plugin names = `excel-mcp` and `excel-cli`
- ✅ Answered #2: Published repo = `sbroenne/mcp-server-excel-plugins`
- ✅ Answered #3: Custom Excel agent = YES for MCP only (agent scaffolds conversational workflows), NO for CLI (scripting needs no conversation)

### Phase 5: User Acceptance (Stefan) — 20260423T174800Z

Stefan approved Kelso's override and refinements. Green-lit:
- MCP binary via GitHub Release download (accept 2-step install)
- Excel agent for MCP plugin only
- Lockstep versioning (plugin version = MCP server release tag)
- Windows-only multi-layered gating
- Automated GitHub Action publication

---

## Final Decisions (ALL LOCKED)

| # | Topic | Decision |
|---|-------|----------|
| 1 | Plugin names | `excel-mcp` (MCP server + agent) and `excel-cli` (CLI skill only) |
| 2 | Published repo | `sbroenne/mcp-server-excel-plugins` (new repo, created during Phase 0) |
| 3 | Custom agent | YES for MCP (thin agent enforcing CRITICAL RULES + workflow hints), NO for CLI |
| 4 | MCP binary distribution | GitHub Release download via `bin/download.ps1` (not Git-committed) |
| 5 | Marketplace submission | DEFER to v2 (not submitting to github/copilot-plugins in v1) |
| 6 | Plugin layout | TWO separate plugins (clean separation, users install only what they need) |
| 7 | Publication mechanism | Automated GitHub Action (deviates from office-coding-agent manual precedent) |

**BONUS DECISIONS (Kelso):**
- Version pinning: Lockstep (plugin version = MCP server release tag)
- Windows gating: Multi-layered (plugin.json description, keywords, SKILL.md preconditions, README warning, runtime graceful failure)

---

## Deliverables

**Final Architecture Document:** `.squad/decisions/inbox/kelso-plugin-shape-final.md`
- Complete two-plugin layout (excel-mcp + excel-cli)
- Published repo structure (`sbroenne/mcp-server-excel-plugins/`)
- Phased execution plan (Phase 0–4, ~9 hours total to "installable")
- GitHub Action automation sketch for cross-repo publish
- Installation workflow examples for both plugins

**Published Repo Structure:**
```
mcp-server-excel-plugins/
├── README.md  (installation, Windows warning)
├── .gitignore (ignore bin/*.exe, keep bin/download.ps1)
└── plugins/
    ├── excel-mcp/
    │   ├── plugin.json
    │   ├── .mcp.json
    │   ├── agents/excel.agent.md
    │   ├── skills/excel-mcp/SKILL.md
    │   └── bin/download.ps1 (committed), mcp-excel.exe (gitignored)
    └── excel-cli/
        ├── plugin.json
        └── skills/excel-cli/SKILL.md
```

---

## Status

✅ **COMPLETE** — All 6 decisions locked, Kelso's plan finalized, Stefan approved, ready for Phase 0 (create published repo).

**Next steps:**
- Scribe: Log decisions to decisions.md, commit .squad artifacts
- Coordinator: Assign Phase 0 (repo creation) + Phase 1–4 (build, test, automation, docs) to implementation team
- rubber-duck (parallel review): Finalize critique, submit findings to Coordinator

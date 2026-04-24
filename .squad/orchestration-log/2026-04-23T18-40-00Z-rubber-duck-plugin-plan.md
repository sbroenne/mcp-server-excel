# Orchestration: Rubber-Duck Review — Plugin Plan

**Date:** 2026-04-23T18:40:00Z  
**Agent:** rubber-duck-plugin-plan  
**Model:** Claude Sonnet 4.5  
**Status:** COMPLETE  
**Verdict:** APPROVE WITH CONDITIONS

---

## Scope

Conducted structured rubber-duck critique of Kelso's Copilot CLI plugin plan for ExcelMcp (Phases 0–4).

---

## Findings Summary

### Critical Issues (4)

1. **Wrapper Script for Missing Binary** — User forgets `download.ps1`, MCP fails with cryptic error
2. **`.mcp.json` `{pluginDir}` Placeholder Unvalidated** — Assumption about CLI placeholder expansion unproven
3. **GitHub Action Uses Over-Permissioned PAT** — Replace with GitHub App scoped to plugin repo
4. **No SHA256 Verification on Binary Download** — MITM/corruption risk

### Moderate Issues (4)

5. Version skew detection (binary version vs. plugin version)
6. Publish workflow atomicity (concurrency control, single commit)
7. CLI plugin discovery without agent presence
8. Custom frontmatter fields not in Copilot CLI spec

### Open Questions (5)

Q1. Does `mcp-server-excel` have release workflow with binary assets?  
Q2. Binary availability race condition with plugin publish?  
Q3. Corporate proxy/firewall support?  
Q4. Air-gapped offline fallback?  
Q5. Does dual install (excel-mcp + excel-cli) download binary twice?

---

## User Decision

**Stefan Brönner approved:** Option (a) — Accept all 4 critical fixes + Phase -1 spike.

---

## Phase -1 Spike (NEW)

**Added to plan by user directive:**
- Validate `{pluginDir}` placeholder expansion before proceeding
- Minimal "hello-world" plugin install test
- Document working mechanism or pivot
- Exit criteria: Working install flow confirmed

**ONLY proceed to Phase 0 if spike succeeds.**

---

## Outcome

✅ Rubber-duck critique complete  
✅ All 4 critical + 4 moderate findings documented  
✅ All 5 open questions answered  
✅ Phase -1 spike authorized  
✅ Plan ready for Kelso refinement

---

## Next Step

Kelso incorporates all findings + spike phase into refined plan document.

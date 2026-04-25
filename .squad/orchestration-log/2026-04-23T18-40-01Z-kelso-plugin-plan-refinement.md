# Orchestration: Kelso Plugin Plan Refinement

**Date:** 2026-04-23T18:40:01Z  
**Agent:** kelso-plugin-plan  
**Model:** Claude Sonnet 4.5  
**Turns:** 3  
**Status:** COMPLETE

---

## Scope

Refined Copilot CLI plugin plan for ExcelMcp based on rubber-duck critique findings and user directives.

---

## Input

- **Rubber-Duck Verdict:** APPROVE WITH CONDITIONS
- **Critical Findings:** 4 (wrapper script, placeholder validation, GitHub App auth, SHA256 verification)
- **Moderate Findings:** 4 (version skew, atomicity, CLI discovery, frontmatter)
- **Open Questions:** 5 (Q1–Q5 all answered)
- **User Directive:** Accept all fixes, add Phase -1 spike before Phase 0

---

## Work Completed

### Turn 1: Initial Refinement
- Incorporated all 4 critical fixes into design
- Answered Q1–Q5 with technical details
- Drafted Phase -1 (Spike) requirements

### Turn 2: Documentation & Decision File
- Updated plan document with all findings
- Created `.squad/decisions/inbox/kelso-spike-first-and-fixes.md` with full decision record
- Documented phase flow: Phase -1 → Phase 0–4

### Turn 3: Final Review & History Update
- Updated `.squad/agents/kelso/history.md` with session summary
- Confirmed spike requirements + exit criteria
- Marked plan ready for execution

---

## Deliverables

1. ✅ Refined plan (5 phases: -1 spike, then 0–4)
2. ✅ Decision document (critical + moderate findings, Q&A, implementation details)
3. ✅ Agent history updated with session context
4. ✅ All 4 critical issues addressed with design patterns
5. ✅ Phase -1 spike fully scoped with exit criteria

---

## Key Design Changes

- **Phase -1 (Spike):** Validate `{pluginDir}` placeholder + wrapper pattern before full implementation
- **Wrapper Script:** `bin/start-mcp.ps1` detects missing binary, version skew, provides user guidance
- **GitHub App:** Replaces PAT in release workflow (Phase 4)
- **SHA256 Verification:** Release workflow produces `checksums.txt`, download script validates
- **Version Tracking:** `version.txt` embedded in plugin, binary-to-plugin skew detection

---

## Outcome

✅ All rubber-duck findings incorporated  
✅ Phase -1 spike authorized and fully designed  
✅ Plan ready for implementation  
✅ Decision record complete and auditable

---

## Next Step

Kelso proceeds to Phase -1 spike (validate install mechanism) before starting Phase 0.

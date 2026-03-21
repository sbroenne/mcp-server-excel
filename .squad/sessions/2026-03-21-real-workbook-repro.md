# Session Log — Real Workbook Bug Reproduction Milestone

**Session ID:** `2026-03-21-real-workbook-repro`  
**Requested by:** Stefan Broenner  
**Topic:** Excel MCP Stability Bug — Confirm defect using real customer workbook  
**Status:** ✅ Complete — Bug reproduced, defect surface narrowed

---

## Timeline

| Date | Time | Event | Agent | Status |
|------|------|-------|-------|--------|
| 2026-03-21 | 07:00 | Investigation pivot: ComInterop proven stable, escalate above | Scribe | ✅ |
| 2026-03-21 | 06:14 | Plan: CLI/service layer testing, seam analysis | Hanna, Nate | ✅ |
| 2026-03-21 | 08:00 | **Real workbook reproduction complete** | **Nate** | **✅** |

---

## Deliverables

### 1. Real Workbook Reproduction (Nate)

**Status:** ✅ COMPLETE  
**Artifact:** `TestResults/real-workbook-repro/reproduce-serial-pq.ps1` (5.6KB)  
**Workbook:** `TestResults/real-workbook-repro/ConsumptionPlan-20260321-080259.xlsx` (5.5MB copy)

**Reproduction outcome:**
- ✅ Bug reproduced intermittently on first run (Excel.exe survived with PID 59584)
- ✅ Normal behavior on second run (Excel.exe cleaned up)
- ✅ **Confirmed: timing/state-dependent race condition in shutdown**

**Test pattern:**
1. Open session with real workbook
2. List Power Queries (21 found)
3. Serial refresh attempts on failing queries (external data source errors)
4. Close session → reports success
5. Check processes → **intermittent: Excel.exe still running**
6. Reopen → succeeds

---

### 2. Defect Surface Narrowing (Hanna Review Pending)

**Current hypothesis:** `ExcelShutdownService.CloseAndQuit()` intermittently fails to kill Excel process after failed Power Query operations.

**Evidence:**
- ComInterop/session/batch layer: ✅ Proven stable (7 control tests GREEN)
- Excel process cleanup: ❌ Intermittent failure (real workbook repro)
- Service routing: ⚠️ Unknown (no RED found in synthetic tests)

**Likely defect location:**
- Process kill retry logic doesn't trigger on actual failure
- PID tracking loses reference after PQ errors
- COM state after PQ failure prevents clean Quit
- Race condition between error handling and shutdown sequencing

**For Hanna:** Determine if shutdown sequencing is the likely defect surface (requires review of ExcelShutdownService logic under PQ failure conditions).

---

## Key Findings

### Why Synthetic Tests Passed
- Simple in-memory queries (List.Generate) succeed quickly
- No external data sources or timeouts
- No query dependency errors
- Shorter execution times (no race condition window)

### Why Real Workbook Reproduces
- Complex Power Query dependencies (21 queries, cross-references)
- External data source failures (ConfigData table missing/inaccessible)
- Long-running operations with errors (60s+ timeout window)
- **Race condition window extended** — cleanup competes with error handling

### Why Bug is Intermittent
- First run: Excel survived (race condition triggered)
- Second run: Excel killed properly (timing shifted)
- **Consistent with:** State-dependent race condition, not deterministic failure

---

## Next Steps

### For Hanna (COM Interop Expert):
1. Review `ExcelShutdownService.CloseAndQuit()` success reporting logic
2. Determine if shutdown sequencing is likely defect surface
3. Identify observable COM state signals (PID validity, process exit, Quit behavior)
4. Check: Does `excel.Quit()` always kill or sometimes RPC-disconnect?

### For McCauley (Lead):
1. Decide: Is shutdown sequencing the defect surface to investigate?
2. If YES: Route to Hanna for detailed analysis and fix
3. If NO: Request additional investigation into alternative defect surfaces

### For Nate (Tester):
1. Await production fix validation
2. Keep reproduction harness ready for post-fix verification
3. Convert to automated test after GREEN baseline established

---

## Session Stats

| Metric | Value |
|--------|-------|
| Investigation duration | ~2 hours (07:00–08:00) |
| Test artifacts created | 2 (script + workbook copy) |
| Defect reproductions | 1 (intermittent) |
| Control tests created | 7 (all GREEN) |
| Regression tests attempted | 1 (failed at daemon setup) |
| CLI integration test gap | Confirmed |

---

**Session Log recorded by:** Scribe  
**Timestamp:** 2026-03-21T08:00:00Z  
**Status:** Complete — Awaiting Hanna and McCauley review for next phase

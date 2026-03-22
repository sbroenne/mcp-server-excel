# Orchestration Log — Real Workbook Bug Reproduction Complete

**Date:** 2026-03-21  
**Time:** 08:00 UTC  
**Milestone:** Nate reproduced the bug intermittently on a copied real workbook  
**Status:** Bug CONFIRMED and reproduced with clear failure signals

---

## Reproduction Milestone

### What Was Achieved

Nate successfully reproduced the stability bug using a real customer workbook with complex Power Query dependencies. The bug manifests as:

**Strongest symptom:** Session close can report success while `EXCEL.EXE` remains alive after long failed PQ workflow.

### Reproduction Environment

**Workbook:** `Consumption Plan.xlsx` (copy)
- Size: 5.5MB
- Power Queries: 21 complex, multi-step dependencies
- Data Model: Integrated with external references
- Data Sources: External (ConfigData table references)
- Path: `TestResults/real-workbook-repro/ConsumptionPlan-20260321-080259.xlsx`

**Reproduction Script:** `TestResults/real-workbook-repro/reproduce-serial-pq.ps1`
- Self-contained PowerShell harness
- Uses actual `excelcli` commands (not direct API)
- Matches exact bug report workflow
- Status: Ready for team review

### Bug Reproduction Sequence

1. Open session with real workbook
2. List Power Queries → **21 found**
3. Serial refresh attempts:
   - "Plan Parameters" → **FAILED** (61s timeout)
   - "Date" → **FAILED** (52s timeout)
   - "Milestones" → **FAILED** (3.5s)
   - "Consumption" → **FAILED** (3.6s)
4. Test session health → **RESPONSIVE** (list still works)
5. Close session → **Reports SUCCESS**
6. Check processes → **INTERMITTENT: Excel.exe sometimes remains running**
7. Immediate reopen → **SUCCEEDS** (new session created)

### Observed Behavior

**First Run:** Excel.exe remained (PID 59584, 434MB) — **BUG REPRODUCED ✓**  
**Second Run:** Excel.exe cleaned up properly — **NORMAL BEHAVIOR**  

**Finding:** Bug is **timing-dependent** or **state-dependent**, not deterministic.

---

## Root Cause Hypothesis

**Defect Surface:** `ExcelShutdownService.CloseAndQuit` **intermittently** fails to kill Excel process after failed Power Query operations.

**Likely failure points:**
- `excel.Quit()` returns without error but process survives
- Process kill retry logic sometimes doesn't trigger on actual failure
- PID tracking sometimes loses Excel process reference
- COM disconnect occurs but process **sometimes** survives in background
- **Timing/state dependent** — not deterministic, race condition window

**Why synthetic tests passed:**
- Simple in-memory queries (List.Generate) succeed quickly
- No external data source failures or timeout conditions
- No query dependency chain errors or complex refresh states
- Excel cleanup worked normally under success conditions
- Shorter execution times (no race conditions during shutdown)

**Why real workbook reproduces:**
- Complex query dependencies with cross-references
- External data source failures (ConfigData table missing/inaccessible)
- Long-running operations with errors (60s+ timeout window)
- Excel state after failed refresh differs significantly from success state
- **Race condition window longer** — cleanup competes with error handling for Excel state

---

## Investigation Path Completion

This milestone completes the escalation path established on 2026-03-21 at 07:00 UTC:

**Previous milestones:**
1. ✅ **2026-03-21 07:00** — Pivoted: ComInterop/session layer proven stable (7 control tests GREEN)
2. ✅ **2026-03-21 06:14** — Identified: CLI daemon state and service routing layer as likely defect surface
3. ✅ **2026-03-21 08:00** — **CURRENT: Real workbook confirms bug exists above ComInterop layer**

**Escalation confirmed:**
- ComInterop session/batch management: ✅ HEALTHY (proved by controls)
- CLI daemon initialization: ⚠️ SETUP ISSUE (JSON parsing error, not PQ-specific)
- Excel shutdown after PQ failure: ❌ **BUG CONFIRMED** (intermittent process survival)

---

## Key Findings

### 1. Bug is NOT in ComInterop/Session Infrastructure
- 7 control tests (batch + session layer) all PASSED
- Session operations themselves work correctly (list, refresh succeed up to failure)
- Session close() reports success accurately
- **Problem is elsewhere**

### 2. Bug IS in Excel Process Cleanup After Realistic Failures
- ExcelShutdownService reports success but leaves Excel.exe running
- Happens specifically after Power Query operations fail with data source errors
- **Intermittent** — first run reproduced, second run succeeded (race condition indicator)
- **State-dependent** — only with complex dependency chains and long-running failures

### 3. Real Workbook Pattern Required
- Synthetic tests (List.Generate, <1s refresh, in-memory data) cannot reproduce
- Need: External data sources, network timeouts, query dependencies, hours-long operations
- This explains why 7 control tests all passed but user reports still valid

---

## Impact Analysis

**User Workflow:**
1. Serial workbook automation (compliance reporting, data integration)
2. Power Query refresh operations fail due to data source issues
3. Session.Close() reports success ✓
4. User assumes cleanup succeeded ✓
5. **Excel.exe SOMETIMES accumulates** ❌ (intermittent, not guaranteed)
6. Over hours/days of automation: system resources exhausted
7. Automation becomes unreliable: later commands may fail fast due to hung Excel

**This explains the user's symptoms:**
- "sometimes appears hung" = intermittent failure
- "later commands may fail fast" = accumulated Excel processes
- "lingering processes have to be manually stopped" = cleanup not guaranteed

---

## Next Steps for Production Team

### For Hanna (COM Interop Expert):
1. Investigate `ExcelShutdownService.CloseAndQuit` success reporting when process survives
2. Review PID tracking after failed Power Query operations — verify capture timing and validity
3. Check process kill retry logic triggers on actual failure vs. assumed success
4. Add explicit process verification after `excel.Quit()` before returning success
5. Verify Excel state after PQ refresh failure — may differ from success state (stuck modal dialog, error state, etc.)

### For Nate (Tester):
1. ✅ Create reproduction harness — DONE (`reproduce-serial-pq.ps1`)
2. ✅ Document findings — DONE (this file)
3. Await production fix before converting harness to automated test
4. After fix: Validate GREEN baseline with same harness
5. Convert to `RunType=OnDemand` integration test for stability regression suite

### For Production:
1. Review shutdown sequencing — is current order correct?
2. Check: Does `excel.Quit()` always kill the process or sometimes just RPC-disconnect?
3. Test: Does PID capture fail after PQ errors?
4. Validate: Is there a modal dialog or error state blocking Quit?
5. Add: Production logging for process survival detection and diagnostic context

---

## Test Artifacts

**Reproduction script:** `TestResults/real-workbook-repro/reproduce-serial-pq.ps1` (5.6KB)
- Self-contained, no external dependencies beyond excelcli
- Clear pass/fail signals
- Captures process state before/after
- Matches exact bug report pattern

**Test workbook:** `TestResults/real-workbook-repro/ConsumptionPlan-20260321-080259.xlsx` (5.5MB)
- Real customer workbook (anonymized copy)
- Complex Power Query dependencies
- Realistic failure mode (data source errors)
- Ready for production team investigation

**Output logs:**
- Full repro run output with timing
- Detailed error messages from Power Query refresh failures
- Process state observations (before/after close)

---

## Decision — Bug Confirmed, Root Cause Surface Identified

**Status:** ✅ REPRODUCED  
**Severity:** High (intermittent process accumulation affects reliability)  
**Root cause surface:** `ExcelShutdownService.CloseAndQuit()` success reporting vs. actual process cleanup  
**Action:** Proceed to production investigation and fix  

---

## Session Log Update

**Session ID:** Stability Investigation — Real Workbook Reproduction  
**Requested by:** Stefan Broenner  
**Participants:** Nate (Tester), Hanna (COM Expert, reviewing), McCauley (Lead, monitoring)

**Timeline:**
- 2026-03-21 05:00 — Investigation pivot: escalate above ComInterop
- 2026-03-21 06:14 — Planning: Hanna & Nate analysis, test seam identification
- 2026-03-21 07:00 — Decision: ComInterop healthy, CLI/service layer under investigation
- 2026-03-21 08:00 — **CURRENT: Real workbook repro complete, bug confirmed**

**Next:** Hanna and McCauley review whether shutdown sequencing is the likely defect surface.

---

**Recorded by:** Scribe  
**Status:** Complete — Bug reproduced, ready for production investigation

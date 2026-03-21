# Excel MCP Bug Report — Long-Running Power Query Automation Can Leave Sessions Wedged

**Date:** March 21, 2026  
**Repo version:** `v1.8.32`  
**OS:** Windows 11 (`10.0.26200.0`)  
**Excel:** `16.0.19822.20104`  
**.NET SDK:** `10.0.201`  
**Client:** `excelcli` from local source tree  
**File(s) affected:** Large `.xlsx` workbooks with multiple Power Query refresh steps

---

## Bug Description

Long-running Power Query automation can still leave Excel automation in a wedged state even after several upstream refresh/deadlock fixes.

The failure pattern is:

- one or more Power Query refresh operations run for a long time
- the command may appear hung or eventually time out/cancel
- subsequent operations on the same session become unusable
- `EXCEL.EXE` and/or `excelcli` may remain running until manually killed
- wrapper automation then appears flaky because the poisoned session or lingering Excel process affects later steps

This looks like a remaining runtime stability defect, not just a single bad workbook.

---

## Component

- [ ] **MCP Server**
- [x] **CLI**
- [x] **Core Library**
- [ ] **Not sure**

---

## Command / Usage

Observed with serial `excelcli` automation against a workbook that has multiple dependent Power Query refreshes.

Representative command sequence:

```powershell
excelcli -q session open .\Workbook.xlsx
excelcli -q range set-values --session <sid> --sheet _Config --range Z2 --values <json>
excelcli -q powerquery refresh --session <sid> --query-name "Plan Parameters"
excelcli -q powerquery refresh --session <sid> --query-name "Date"
excelcli -q powerquery refresh --session <sid> --query-name "Milestones"
excelcli -q powerquery refresh --session <sid> --query-name "Consumption"
excelcli -q session close --session <sid>
```

Also observed indirectly through wrapper scripts that perform the same kind of serial workbook sync/refresh workflow.

---

## Expected Behavior

- Long-running Power Query refreshes should either complete successfully or fail cleanly.
- If an operation times out or is cancelled, the session should be closed or made recoverable in a predictable way.
- A failed refresh should not poison later operations or leave `EXCEL.EXE` / `excelcli` lingering indefinitely.
- New session creation should not be blocked by stale dead sessions or half-dead Excel processes.

---

## Actual Behavior

- Long-running refreshes sometimes appear hung for several minutes.
- After a timeout/cancellation, the session may become effectively unusable until Excel is killed manually.
- Later commands may fail fast with timeout-style errors or appear stuck behind a poisoned STA thread.
- Wrapper automation may look nondeterministic because the real failure happened earlier and left the process/session in a bad state.
- Lingering `EXCEL.EXE` and `excelcli` processes have to be manually stopped to recover cleanly.

---

## Error Messages

Representative failure strings seen while investigating:

```text
TimeoutException: A previous operation timed out or was cancelled for '<workbook>'. The Excel COM thread may be unresponsive. Please close this session and create a new one.
```

```text
Excel operation timed out after <N> seconds for '<workbook>'. Excel may be unresponsive or the operation is taking longer than expected.
```

```text
Session for '<workbook>' was disposed while submitting an operation.
```

In practice, the more visible symptom is often not the exception text but the wedged automation state and lingering Excel processes.

---

## Steps to Reproduce

1. Use a workbook with several dependent Power Query refreshes, including at least one long-running refresh.
2. Open it with `excelcli`.
3. Apply any needed config cell update.
4. Run multiple `powerquery refresh` operations in a realistic serial sequence.
5. If one operation runs long enough to hit cancellation/timeout or enters a bad refresh state, continue issuing follow-up operations on the same session.
6. Observe that:
   - the session may become poisoned
   - later operations fail or hang
   - `EXCEL.EXE` / `excelcli` may remain alive until manually killed

This becomes easier to trigger with larger Power Query workloads and wrapper automation that opens/closes many workbooks in one run.

---

## Likely Source Hotspots

### 1. Session timeout poisoning in `ExcelBatch.Execute()`

`src\ExcelMcp.ComInterop\Session\ExcelBatch.cs:471-577`

- `_operationTimedOut` permanently poisons the session after a timeout/cancel
- later operations fail fast even if they were not the original cause
- the STA thread may still be stuck in COM while the session is already considered unusable

### 2. High-CPU polling during refresh completion

`src\ExcelMcp.Core\Commands\PowerQuery\PowerQueryCommands.Helpers.cs:318-367`

- `WaitForRefreshCompletion()` polls `.Refreshing`
- comments explicitly acknowledge CPU-spin trade-offs during large refreshes
- this may contribute to “looks hung / system thrashes / user kills Excel” outcomes

### 3. Lazy dead-session cleanup

`src\ExcelMcp.ComInterop\Session\SessionManager.cs:284-303, 551-565`

- dead Excel processes are cleaned up only when certain accessors are called
- stale sessions can survive long enough to block reopen/recovery paths

### 4. Session creation/open contention

`src\ExcelMcp.ComInterop\Session\ExcelSession.cs:18, 121-159`

- serialized creation lock and timeout behavior may amplify recovery failures after earlier instability

---

## Relevant Upstream History

Several recent releases already targeted the same problem family:

- `v1.8.20` / `#525` — robust refresh + cancellation changes
- `v1.8.21` / `#526` — CPU spin and session hang mitigation
- `v1.8.28` / `#542` — `EnterLongOperation` deadlock fix
- `v1.8.29` / `#543` — cancellation recovery and refresh deadlocks
- `v1.8.30` / `#545` — remaining synchronous COM refresh stabilization

This report is about the remaining stability gap after those fixes, not about the already-fixed deadlock itself.

---

## Suggested Investigation Direction

1. Build one deterministic serial Power Query repro inside this repo's own test space.
2. Run that repro across `v1.8.20`, `v1.8.21`, `v1.8.28`, `v1.8.29`, `v1.8.30`, and `v1.8.32`.
3. Confirm whether the dominant remaining failure is:
   - session poisoning after timeout/cancel
   - refresh polling / CPU-spin side effects
   - dead-session cleanup lag
   - session creation/open contention after partial failure
4. Fix one upstream defect at a time with regression coverage.

---

## Excel Process Cleanup

- [ ] Excel processes clean up properly after the command
- [x] Excel processes remain running (this is part of the bug)
- [ ] Not applicable / unsure

---

*Report generated locally from direct source review plus repeated long-running Power Query automation failures observed against `mcp-server-excel` v1.8.32.*

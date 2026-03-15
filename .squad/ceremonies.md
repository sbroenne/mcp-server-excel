# Ceremonies

> Team meetings that happen before or after work. Each squad configures their own.

## Design Review

| Field | Value |
|-------|-------|
| **Trigger** | auto |
| **When** | before |
| **Condition** | multi-agent task involving 2+ agents modifying shared systems |
| **Facilitator** | lead |
| **Participants** | all-relevant |
| **Time budget** | focused |
| **Enabled** | ✅ yes |

**Agenda:**
1. Review the task and requirements
2. Agree on interfaces and contracts between components
3. Identify risks and edge cases
4. Assign action items

---

## Retrospective

| Field | Value |
|-------|-------|
| **Trigger** | auto |
| **When** | after |
| **Condition** | build failure, test failure, or reviewer rejection |
| **Facilitator** | lead |
| **Participants** | all-involved |
| **Time budget** | focused |
| **Enabled** | ✅ yes |

**Agenda:**
1. What happened? (facts only)
2. Root cause analysis
3. What should change?
4. Action items for next iteration

---

## COM Interop Review

| Field | Value |
|-------|-------|
| **Trigger** | auto |
| **When** | after |
| **Condition** | any change touching COM interop calls, Excel Object Model usage, or `src/ExcelMcp.ComInterop/` |
| **Facilitator** | Hanna |
| **Participants** | Hanna + author of change |
| **Time budget** | focused |
| **Enabled** | ✅ yes |

**Agenda:**
1. Verify every COM call against Excel Object Model documentation
2. Check COM object lifecycle (acquire in try, release in finally)
3. Validate property types (double→int conversions, date handling)
4. Confirm collection indexing is 1-based
5. Flag any undocumented API usage — must cite docs or reject

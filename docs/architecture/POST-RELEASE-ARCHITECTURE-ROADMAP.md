# Post-release architecture roadmap

This roadmap sequences larger changes after the current release. It is an implementation contract, not a promise to broaden Excel automation beyond what COM/STA and MCP can safely guarantee. Each phase ships behind compatibility-preserving adapters and is measured against a fixed real-Excel benchmark corpus.

## Constraints that shape every phase

- MCP `tools/list` is authoritative for the public surface. Tool schemas and result shapes consume client context, so additions must be deliberate and pagination/list-change behavior must remain compatible.
- Mutations for one session serialize. A client approval is a pause owned by the client; the server must not sleep, poll, or auto-approve while waiting.
- COM calls are non-preemptible and cancellation is cooperative. If dispatch may have happened and the outcome is not observed, the receipt is `unknown` and the mutation is not replayable automatically.
- Keep the existing operation names and behavior as a compatibility layer until replacement coverage and telemetry are proven. No phase implies generic workbook transactions, rollback, or exactly-once semantics.

## Ranked work

### 1. Bounded large-read `ResultProjection` (first)

Introduce a shared projection envelope for read results: `preview`/`full` mode, maximum rows, columns, and bytes, plus `totals`, `truncated`, `nextRange`, and a stable content `fingerprint`. Keep small reads byte-for-byte compatible where practical; make truncation explicit rather than silently dropping cells.

**Impact:** predictable token and memory use; clients can page or request a targeted follow-up; one result contract feeds range, table, and inspection tools. **Risk:** callers may mistake a preview for complete data, or fingerprints may be expensive/stale. **Effort:** medium. **Dependencies:** central range inspection and serialization helpers; benchmark corpus; schema/version policy.

**Acceptance/benchmarks:** for configured limits, no response exceeds the byte cap (excluding protocol framing); a 1M-cell fixture returns within the memory budget and marks `truncated=true`; `nextRange` requests cover every omitted rectangle without overlap; repeated unchanged reads produce the same fingerprint; preview/full parity is verified on representative values, formulas, errors, and empty cells.

### 2. Generated declarative `CommandCatalog` (second)

Define command metadata once (name, read/write safety, dispatcher, plan eligibility, profile/action cost, schema, result projection, and approval requirements) and generate the MCP tool listing, safety checks, dispatcher registration, plan validation, profiles, and result metadata from it.

**Impact:** eliminates drift between five policy surfaces and makes capability review auditable. **Risk:** generator defects can remove or over-permit tools; generated diffs may be noisy. **Effort:** medium-high. **Dependencies:** stable metadata types from phase 1; golden generated artifacts; CI validation that generated output is current.

**Acceptance/benchmarks:** every public operation has exactly one catalog entry; generated `tools/list`, safety, dispatcher, and plan/profile metadata agree in a machine comparison; a catalog change produces a reviewable diff and fails CI if artifacts are stale; startup registration time does not regress by more than 10% on the benchmark host.

### 3. Generalized safe `PlanExecutor` (third)

Replace command-specific batching with a plan executor that performs one validation pass, one approval/checkpoint boundary, and one STA dispatch per plan, returning a receipt for every step. Preserve operation order and mark unknown outcomes after a possible COM dispatch; never retry an unknown mutation.

**Impact:** consistent multi-step safety, lower dispatcher overhead, and durable explainability. **Risk:** a plan can enlarge blast radius; validation may become a bottleneck; partial completion is unavoidable. **Effort:** high. **Dependencies:** CommandCatalog; session serialization; checkpoint/receipt format; client pause protocol.

**Acceptance/benchmarks:** mixed read/write plans reject invalid or disallowed steps before dispatch; same-session plans never overlap; each step has status, timing, and outcome (`applied`, `rejected`, `failed`, or `unknown`); a fault-injected timeout demonstrates no automatic replay; 100-step representative plans reduce STA dispatches by at least 30% versus the current baseline without changing observable ordering.

### 4. Action-level token-budgeted MCP profiles (fourth)

Add profiles that select catalog actions and result defaults by task (for example, inspect, edit, or audit), with explicit token/byte budgets and a safe fallback. Profiles constrain discovery and response size; they do not bypass safety or approval.

**Impact:** smaller context and more predictable latency for Copilot-style clients. **Risk:** a profile can hide a needed action or become stale as commands evolve. **Effort:** medium. **Dependencies:** generated catalog and `ResultProjection`; tools/list pagination and list-change signaling.

**Acceptance/benchmarks:** each profile has a documented action allowlist and budget; generated profile membership matches catalog metadata; representative Copilot conversations stay within the configured result budget; an unavailable action yields a discoverable, actionable error and no mutation.

### 5. Per-session `SessionSlot` (fifth)

Consolidate fragmented per-session dictionaries and 10 ms polling into one `SessionSlot` owning workbook identity, STA queue, lock, lifecycle, receipts, and health state. Use awaitable signaling and bounded admission for mutations; retain one-owner-per-session semantics.

**Impact:** fewer races/leaks and lower idle CPU; clearer cleanup and observability. **Risk:** migration can strand existing sessions or alter timeout behavior. **Effort:** high. **Dependencies:** PlanExecutor receipts; host identity fields; compatibility adapter for current session lookup.

**Acceptance/benchmarks:** no periodic polling remains on the session hot path; idle CPU is below 1% per session in a 5-minute soak; concurrent same-session mutations execute in submission order; close/reopen tests release all COM references; bounded queue behavior is measured (wait/reject, never drop a write).

### 6. Host-owned runtime identity, catalog, and profile hashes (sixth)

Have the host publish immutable runtime identity (server/build, Excel PID plus start time where available), and hashes for the active command catalog and profiles. Include them in capabilities and receipts so clients can detect drift and support cases can reproduce a run.

**Impact:** trustworthy diagnostics and cache invalidation. **Risk:** exposing unstable or sensitive host details; hash changes can cause unnecessary client refreshes. **Effort:** low-medium. **Dependencies:** catalog/profile generation; SessionSlot lifecycle; redaction policy.

**Acceptance/benchmarks:** every receipt and capability snapshot carries the same catalog/profile hash for its session; a changed catalog causes a list-change signal or explicit refresh guidance; PID reuse tests reject stale identities; hashes are deterministic across identical builds.

### 7. True async bridge (later, exploratory)

Only after the preceding phases and measured bottlenecks, evaluate an asynchronous bridge that separates request waiting from STA execution while preserving per-session ordering and unknown-outcome semantics. Do not claim that async waiting makes COM cancellable.

**Impact:** potentially better multiplexing and client responsiveness. **Risk:** highest; added queues and failure modes can obscure mutation state. **Effort:** very high. **Dependencies:** all prior receipt, identity, admission, and telemetry work; a fault/latency test harness; explicit product decision on process isolation.

**Acceptance/benchmarks:** a prototype shows improved concurrent read latency without overlapping same-session mutations, bounded memory under load, and correct `unknown` receipts under forced hangs. Proceed only if the measured benefit exceeds operational complexity; otherwise keep the current STA bridge.

## Migration order and release gates

1. Land projection types and adapters; run parity, size, and paging benchmarks.
2. Generate the catalog while retaining the current registry; compare outputs in CI.
3. Introduce PlanExecutor for opt-in plans, then migrate existing batch paths after fault-injection and approval-pause tests.
4. Publish profiles and budget telemetry; update Copilot-facing schemas only after `tools/list` compatibility review.
5. Move session state behind `SessionSlot`, with dual-read telemetry and a rollback switch until soak tests pass.
6. Add host hashes and identity to diagnostics/receipts; document redaction and support procedures.
7. Reassess the async bridge from production measurements; it is not a prerequisite for the earlier phases.

Each gate requires integration tests against real Excel, a benchmark report (latency, bytes, memory, queue depth, and unknown outcomes), and a compatibility note for existing MCP clients.

## Explicit non-goals

- No generic transaction, rollback, exactly-once, or safe automatic retry guarantee for Excel mutations.
- No dropping queued writes under load, reordering observable operations, or suppressing required Excel events.
- No server-side waiting for human approval and no assumption that clients refresh `tools/list` without notification/support.
- No warm multi-tenant Excel pool, distributed unattended automation platform, or formula-engine parity claim.
- No source, test, or existing documentation restructuring as part of this roadmap.

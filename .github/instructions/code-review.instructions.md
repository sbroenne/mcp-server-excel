---
applyTo: "**"
excludeAgent: "cloud-agent"
---

# Copilot Code Review

Report only actionable, high-confidence defects introduced by the pull request.
Prioritize correctness, data loss, resource leaks, deadlocks, broken contracts,
and security. Do not request style-only changes, speculative refactors, or
unrelated cleanup.

## Interop and Resource Safety

- Enforce PIA-first interop across Core and ComInterop. Flag new `dynamic`
  Excel access when `Microsoft.Office.Interop.Excel` exposes a strongly typed
  member. Allow late binding only for a documented PIA or runtime dependency
  gap, such as `Application.Run`, `AutomationSecurity`, or `VBProject`; do not
  suggest early binding that reintroduces Office.Core or VBE dependencies.
- Flag direct numeric casts from `dynamic` COM values. Excel can marshal integer
  properties as `double`; use `Convert.ToInt32`, `Convert.ToDouble`, or the
  equivalent explicit conversion instead of `(int)` or an enum cast.
- Verify every acquired COM object is released in a `finally` block. Cleanup
  must target resources by tracked identity, such as PID and start time or a
  stable object identifier, never by process name or display-name substring.
- Do not treat a generic HRESULT, especially `0x800A03EC`, as proof of one
  specific cause without an additional precondition or distinguishing check.
  Route known transient COM failures through the established retry path; do not
  add broad catches inside Core commands instead of `batch.Execute` propagation.
- Verify operation counters, flags, locks, and session state are restored on
  every exit path. A caught timeout or connection failure must not become an
  empty successful result, and one action failure must not terminate the whole
  MCP session or strand Excel.

## Contract and Path Completeness

- When an action or parameter is added, renamed, retyped, or given a new
  default, trace it through the Core interface, CLI argument routing, batch JSON
  dispatch, generated service dispatch, MCP schema, and both entry points.
  Names, types, aliases, defaults, validation, and timeout behavior must agree;
  unknown values must fail explicitly rather than silently select a default.
- When a change adds a dependency, response field, wait, lock, or synchronization
  path, verify the matching existence guard, integrity check, return-value
  check, timeout handling, and round-trip assertion were extended with it.
- When a pull request fixes a value, state, cache, or measurement bug, inspect
  fallback, retry, degenerate, cached-state, and parallel branches for the same
  defect. A primary-path fix is incomplete if another branch still uses the old
  value or behavior.
- If one generated, templated, or intentionally parallel artifact changes,
  verify its source of truth and counterparts change consistently. Flag
  hand-maintained duplicate logic and validation that compares generated output
  against the same stale fallback used to produce it.

## Protocol, Tests, and User-Facing Surfaces

- Reserve MCP stdio stdout exclusively for JSON-RPC. Any reachable log, banner,
  diagnostic, installer, or bootstrap output must go to stderr, and protocol
  responses must be flushed without waiting for process exit.
- After renames or behavior changes, search for stale names, flags, defaults,
  XML summaries, help text, banners, logs, skills, and documentation. Flag
  user-facing counts or versions copied as literals when they can be derived
  from the authoritative source.
- Tests must verify the resulting Excel state and every newly returned field,
  not only `Success`. Check that collection, byte, and encoding assertions use
  the intended xUnit overload and cannot index beyond the final element.

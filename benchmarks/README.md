# Excel MCP baseline benchmark suite

This opt-in suite establishes repeatable before/after evidence for the first ten improvement plans. It exercises four public seams: real Excel behavior, the service boundary, the in-memory MCP protocol boundary, and pure statistical/reporting logic.

The benchmark projects are deliberately not part of the normal solution or CI path. Running the ordinary repository tests does not launch Excel or run performance workloads.

## What a run produces

Each run writes three source-of-truth files:

- `baseline.json`: run state, environment, configuration, every observation (including failed safety observations), invariant evidence, per-case median/p95/p99, bootstrap median confidence intervals, and reliability bounds.
- `observations.csv`: flat raw data for Excel, R, Python, or another statistics tool.
- `baseline.md`: a compact human-readable result.

The JSON and CSV are the comparison evidence. The rendered workbook is a presentation layer, not a replacement for raw data.

The reporter rewrites all three files after every completed scenario with `runState: in-progress`, then writes `runState: completed` only after the entire requested plan set finishes. This preserves earlier results when a later real-Excel scenario hangs or fails. The comparator rejects any in-progress report.

## Profiles

| Profile | Warmups | Timed repetitions | Fault/reliability repetitions | Use |
|---|---:|---:|---:|---|
| `quick` | 1 | 3 | 3 | Smoke-check the harness and get rough initial figures. |
| `standard` | 3 | 10 | 20 | Day-to-day before/after development comparison. |
| `reliable` | 5 | 30 | 100 | Release decision; use on a quiet, stable machine. |

For timing claims, use at least `standard`. Treat `quick` numbers as preliminary. Zero failures in 100 trials still does not prove zero risk: the rule-of-three upper 95% failure-rate bound is about 3%.

## Run it

Close unrelated workbooks and background-heavy applications. Keep power mode, Excel version, build configuration, and machine unchanged between baseline and candidate runs.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Run-BaselineBenchmarks.ps1 -Profile standard -ShowExcel
```

The process-scoped `Bypass` form works on machines that block direct `.ps1` execution and does not change the machine-wide PowerShell policy.

Run only selected plans while developing:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Run-BaselineBenchmarks.ps1 -Profile quick -Plans '04,05,06' -ShowExcel
```

Validate the benchmark/reporting logic without running the real-Excel workloads:

```powershell
dotnet test .\benchmarks\ExcelMcp.Benchmarks.Tests\ExcelMcp.Benchmarks.Tests.csproj `
  --configuration Release `
  -p:ExcelMcpSkipSkillGeneration=true
```

The skip property applies only to the MCP server's release-time skill-document generation. The benchmark still builds and exercises the same server code; avoiding that unrelated build task prevents its loaded assembly from locking a subsequent benchmark build.

The harness creates unique temporary workbooks, owns only the Excel processes it starts, and never terminates pre-existing Excel processes. `-ShowExcel` makes those test instances visible; omit it for lower-noise timings.

Exit code `0` means every acceptance invariant passed. Exit code `2` means the run completed and wrote reports but one or more planned capabilities or safety checks are still red. Exit code `1` is a harness/runtime failure; `3` is cancellation or maximum-duration expiry.

## Compare a candidate one-to-one

Run the candidate with the same profile, plan set, and machine conditions, then:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Compare-Benchmarks.ps1 `
  -Baseline .\artifacts\benchmarks\baseline\baseline.json `
  -Candidate .\artifacts\benchmarks\candidate\baseline.json `
  -OutputDirectory .\artifacts\benchmarks\comparison
```

The comparator refuses non-equivalent evidence. Both reports must be completed and must have the same profile, configuration, Excel visibility, repetition counts, machine hash, Excel version, OS, architecture, logical processor count, scenarios, cases, and metrics. A speed improvement is never accepted if a candidate safety invariant fails or a measured reliability case fails.

For a high-confidence experiment, alternate run order (`baseline, candidate, candidate, baseline`) or randomize it, repeat on at least two fresh machine boots, and compare paired medians and p95s. Keep both raw reports; do not compare a cold first run with a warmed second run. Case-level distributions are the decision surface because scenario-wide summaries can mix different workload sizes.

## Token efficiency

The MCP probe captures exact UTF-8 request/response bytes through the real in-memory MCP transport and stores response hashes. `token_estimate` is deterministically `ceil(bytes / 4)`. It is intentionally labeled an estimate because exact tokens depend on a particular model and tokenizer. Use exact byte changes for the strongest one-to-one protocol comparison.

## Safety limits of the baseline

- A direct COM edit stands in for a human/manual external edit. VBA-specific invalidation is documented but not run because macro trust policy varies by machine.
- Crash durability currently uses disposal without `session.close`; it is a crash-like restart, not a power-cut/fsync proof. A future implementation needs injectable transition-phase crash hooks.
- PID reuse cannot be forced safely through the public seam. Plan 09 therefore keeps `identity_mismatch_fails_closed` red until a PID/start-time/identity seam exists.
- Plan 07 uses the existing review ID as a retry surrogate and records `idempotency_key_supported = 0`; the receipt-replay acceptance check stays red until a true key contract exists.
- Queue depth is inferred from load, latency, rejection count, and working-set growth because there is no public queue-depth counter.
- Plan 10 reports three explicit public-MCP cases on every repetition: legacy calls, `workflow.execute-plan` with legacy open/describe, and `workflow.open-and-describe` plus `workflow.execute-plan`. Every case uses a fresh client, captures initialize, one `tools/list`, and bidirectional `tools/call` bytes, verifies values through MCP, and closes with `save:false`; it does not auto-probe or retry a workflow capability.

See [BENCHMARK-MATRIX.md](BENCHMARK-MATRIX.md) for the exact workload and gate for each plan.

# Anthropic code-execution-with-MCP analysis — 2026-08-06

## Scope and evidence

This note compares the checked-out `feature/safe-compatible-automation` revision `b30ca8e922c933c2709be213c7cda9296ae9364e` with Anthropic's primary engineering article, *Code execution with MCP: Building more efficient agents* (published 2025-11-04; canonical URL: [anthropic.com/engineering/code-execution-with-mcp](https://www.anthropic.com/engineering/code-execution-with-mcp)). The supplied local archive was read at `C:\Users\Ross\Downloads\Code execution with MCP_ building more efficient AI agents _ Anthropic.htm`. Repository claims below are grounded in the checked-out source, tests, and benchmark harness only; no secondary sources were used.

Anthropic's central proposal is **not** a new MCP server feature: the agent receives a code-execution environment, discovers MCP tools as code APIs on demand, runs composition/filtering there, and returns only selected output to the model. This makes the client/execution harness—not the MCP server—the main architectural addition.

## What the Anthropic article actually claims

- Direct tool calling can consume context twice: tool schemas are loaded up front, and a large intermediate result is placed in model context before it is passed to the next tool. The article's example says a two-hour transcript can add 50,000 tokens; it also warns that large documents can exceed context limits. [Anthropic article](https://www.anthropic.com/engineering/code-execution-with-mcp)
- Its proposed harness exposes connected servers as importable functions/files. The agent lists a server directory, reads only the selected tool definitions, then writes code that calls those functions. In its illustrative 150,000-to-2,000-token scenario, Anthropic reports a 98.7% reduction. This is an **article-specific example**, not a result demonstrated by this repository. [Anthropic article](https://www.anthropic.com/engineering/code-execution-with-mcp)
- The harness can filter/aggregate data, iterate, branch, wait, and handle errors before logging a compact result; it can also persist intermediate files and reusable skills. It proposes client-side PII tokenization/de-tokenization and deterministic data-flow rules so raw data can move between tools without entering the model context. [Anthropic article](https://www.anthropic.com/engineering/code-execution-with-mcp)
- Anthropic explicitly treats sandboxing, resource limits, and monitoring as required operational costs of running agent-generated code. The lower-token/lower-latency benefits must be weighed against those costs. [Anthropic article](https://www.anthropic.com/engineering/code-execution-with-mcp)

## Comparison with the current implementation

| Article principle | Current evidence | Assessment |
| --- | --- | --- |
| Reduce schema load | The server has a registration-time `full` or nine-tool `copilot-compact` profile. The compact instructions tell the client that live `tools/list` is authoritative and it can restart with `full` for omitted tools. Profiles govern discovery/context only, not authorization. [Profile catalog](../../src/ExcelMcp.McpServer/McpToolProfile.cs#L10-L72) | **Partly implemented.** This reduces the initial surface, but it is a fixed profile rather than per-task, progressive tool discovery or detail-level search. |
| Collapse round trips | `workflow(open-and-describe)` returns a bounded workbook summary, and `workflow(execute-plan)` submits ordered service commands with an optional bounded verification receipt. A public real-Excel test proves a two-write plan can run in one STA dispatch. [Workflow surface](../../src/ExcelMcp.McpServer/Tools/ExcelWorkflowTool.cs#L45-L159) [Public transport test](../../tests/ExcelMcp.McpServer.Tests/Integration/Tools/WorkflowToolRealExcelTests.cs#L41-L137) | **Implemented at the server layer.** It is composition of declared Excel commands, not arbitrary agent code calling MCP APIs. |
| Keep large results out of context | Verification is caller-selected and capped at 10,000 inspected cells with a 2×4 preview, fingerprint, and explicit partial-verification status. [Verifier](../../src/ExcelMcp.Service/Workflow/WorkflowRangeVerifier.cs#L12-L22) [Large-range test](../../tests/ExcelMcp.McpServer.Tests/Integration/Tools/WorkflowToolRealExcelTests.cs#L141-L185) | **Partly implemented.** Compact receipts prevent some unnecessary read-backs, but ordinary data-returning tools are still direct MCP results; there is no general execution-side filter/transform layer. |
| Reuse code and state | The project provides skill guidance, persistent CLI sessions, workflow capabilities, journals, checkpoints, and recovery. [CLI overview](../../src/ExcelMcp.CLI/README.md#L1-L24) [Safety workflow](../../README.md#L59-L65) | **Adjacent capability.** Session/recovery state is present, but there is no agent workspace where generated code/functions become managed reusable skills. |
| Preserve sensitive intermediate data | Safety controls protect mutation/recovery workflow, but no current source found an execution-harness data-flow policy or a PII tokenization/untokenization interceptor. | **Not implemented.** This is a searched-source absence, not evidence that direct MCP results are safe to disclose. |
| Run agent-written code safely | ExcelMcp drives desktop Excel as the local user; it explicitly excludes server-side and high-volume batch use. [Scope](../../README.md#L99-L110) | **Not implemented—and appropriately not implied.** Adding code execution would create a distinct sandboxed product boundary, not a small extension of `execute-plan`. |

The CLI is already a useful alternative to the article's *context* objective: it wraps shared operations in one tool with skill-based guidance and documents a 64% lower token figure than the traditional MCP surface. However, the project correctly cautions that client behavior determines which schemas actually enter model context. [CLI README](../../src/ExcelMcp.CLI/README.md#L1-L18) [Root README](../../README.md#L125-L170)

## Already implemented ideas worth retaining

1. **Compact, capability-described surfaces.** The runtime returns the active profile, version, tool list, and manifest hash, so a client can discover exactly which optimized surface it received. [Runtime identity](../../src/ExcelMcp.Service/Workflow/WorkflowRuntimeIdentity.cs#L7-L119)
2. **Server-side composition with conservative semantics.** `execute-plan` preserves ordered execution, offers an idempotency key, bounded verification, one checkpoint option, and a gated fast path. The fast path accepts only a narrow compatible command set and falls back before dispatch otherwise. [Workflow fast-path policy](../../src/ExcelMcp.Service/Workflow/WorkflowFastPathPolicy.cs#L7-L80)
3. **Measured protocol footprint.** The benchmark captures real in-memory MCP UTF-8 request/response bytes and labels `ceil(bytes / 4)` as a deterministic estimate rather than model-token truth. [Benchmark method](../../benchmarks/README.md#L74-L76)
4. **Safety before speed.** The comparator requires matched environments and rejects a speed claim if a safety or measured reliability invariant fails. [Comparison rules](../../benchmarks/README.md#L59-L72)

## Methodology changes

Do not import Anthropic's 98.7% claim into project documentation or acceptance criteria. Its number depends on a hypothetical tool population, chosen schemas, execution harness, model, and task. Instead:

1. Treat exact wire bytes as the common primary metric, and collect target-client/model tokenizer counts separately.
2. Compare `full`, `copilot-compact`, CLI, and workflow variants on the *same* task corpus; record discovery bytes, tool-call bytes, model-input tokens when available, elapsed time, completion correctness, and Excel safety outcome.
3. Keep the repository's existing controls: fresh clients for each public-MCP case, identical machine/Excel/configuration, raw observations, medians plus p95/p99, bootstrap intervals, and safety gates before ranking. [Benchmark matrix](../../benchmarks/BENCHMARK-MATRIX.md#L8-L20)
4. Add an execution-harness arm only if one is built. Attribute savings separately to (a) compact registration, (b) workflow batching, and (c) execution-side filtering; otherwise the result cannot establish which design caused the change.

## New opportunities

1. **Demand-driven, compatibility-preserving discovery.** Add an opt-in `search-tools`/manifest route with name-only, description, and full-schema detail levels, while retaining stable full and compact profiles for existing clients. This is the closest direct application of the article without executing arbitrary code.
2. **Typed, server-side data operations.** Offer narrowly scoped filtering, projection, aggregation, join, and bounded-export actions for workbook data. These can keep large values out of the model while preserving ExcelMcp's existing command/safety model; do not first expose arbitrary local scripts as a shortcut.
3. **Plan review at the workflow boundary.** Current documentation is explicit that `execute-plan` is a transport optimization, not a transaction or proof of human approval. A plan-level approval receipt would better match multi-step automation than only per-command review. [Workflow limitation](../../FEATURES.md#L48-L55)
4. **A separate execution-harness prototype, if product demand warrants it.** Expose a small allowlisted API rather than the full user filesystem/COM surface, give it a per-session capability token and isolated workspace, cap time/memory/output, redact logs, and make egress/data-flow policy explicit. This should be a separate security-reviewed mode.

## Risks and boundaries

- **Do not conflate batching with transactions.** Earlier successful plan steps are not rolled back; timeout/lost-connection outcomes must never be replayed automatically. [Workflow limitation](../../FEATURES.md#L48-L55)
- **Do not overstate verification.** A large range can be `partiallyVerified`; the tested 10,100-cell case inspects 9,999 cells and labels the limitation. [Large-range test](../../tests/ExcelMcp.McpServer.Tests/Integration/Tools/WorkflowToolRealExcelTests.cs#L141-L185)
- **Code execution broadens the trust boundary.** This is an inference from the article's required sandboxing and from ExcelMcp running desktop automation locally: an execution prototype must protect against arbitrary file/process/network access, resource exhaustion, secret leakage through logs, and unsafe Excel side effects.
- **Benchmark limits remain explicit.** The current suite does not claim a power-cut durability proof, cannot safely force PID reuse, lacks a public queue-depth counter, and says its public-MCP measurement is not Copilot prompt packing or tokenizer behavior. [Known limits](../../benchmarks/README.md#L78-L85)

## Recommended experiments

| Priority | Experiment | Success evidence / stop condition |
| --- | --- | --- |
| P0 | Run the existing public-MCP protocol probe across `full` and `copilot-compact`, then add real target-client/model token captures. | Publish raw wire bytes and model tokens separately; reject any faster candidate that fails correctness or safety gates. |
| P0 | Compare legacy calls, `open-and-describe` + `execute-plan`, and CLI on a fixed workbook corpus with 1, 8, and 64 compatible operations. | Report discovery bytes, calls, latency, p95, completion, and STA dispatches; include incompatible fast-path fallbacks. |
| P1 | Add sparse mutations outside the verifier's top-left sampled area and ranges over 10,000 cells. | Quantify what a partial receipt can and cannot establish before changing its user-facing wording or adding multi-region/full-hash options. |
| P1 | Prototype plan-level review/approval receipts without arbitrary code execution. | Verify rejection before mutation, receipt binding to exact arguments/session/workbook state, expiry, and unknown-outcome handling. |
| P2 | Only after the above, build a throwaway isolated execution-harness spike around an allowlisted read/filter/export API. | Demonstrate resource/output caps, denied filesystem/network/process access, redacted logs, policy-enforced data flow, and no regression in Excel safety tests; stop if those controls cannot be made observable and testable. |

The recommended near-term direction is therefore **measure and deepen the existing compact/workflow design first**. It captures much of the article's round-trip and result-minimization value without turning a local Excel automation bridge into a general agent-code runtime.

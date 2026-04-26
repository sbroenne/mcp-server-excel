# MCP/CLI Parity Triage

Use this when a bug report says "works in CLI, fails in MCP" or the reverse.

## Workflow

1. Start from the issue payload and identify the exact tool/action pair.
2. Run the narrowest existing MCP regression bucket for that payload.
3. Run the matching CLI parity bucket sequentially (not in parallel with another build/test run).
4. Interpret the split:
   - **MCP fails, CLI passes:** inspect MCP schema, hand-written tool routing, service forwarding, and response envelope handling.
   - **CLI fails, MCP passes:** inspect CLI command wiring, daemon/service transport, and exit-code/error serialization.
   - **Both pass:** treat the issue as already fixed on the branch; do not speculate. Ask for or capture a fresh repro on the current build.
5. Only change code after a red focused test proves the bug still exists.

## Why it works

This isolates transport-layer parity bugs from shared Core logic quickly. It also prevents re-fixing closed issues when current HEAD is already green.

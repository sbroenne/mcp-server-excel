# ExcelMcp review priorities

These instructions apply to code review, not implementation tasks.
Coding instructions are maintained in [AGENTS.md](../AGENTS.md).

Report high-confidence defects introduced by the PR, not style or unrelated
cleanup. Retain these independent checks:

- Typed PIAs first, except documented runtime gaps (`Application.Run`, VBE,
  Office-core). Do not reintroduce unavailable dependencies. Dynamic COM numeric
  values need `Convert.*`, not direct casts.
- Acquired COM references require `finally` cleanup, excluding session-owned
  objects. Process cleanup uses PID/start-time identity, never process names.
- `0x800A03EC` is not a unique diagnosis. Preserve batch/Service error context,
  cancellation, and session recovery; a timeout must not become success.
- Core contracts must agree across generated Service, CLI options/batch JSON,
  MCP schemas, and manual tool exceptions, including defaults and timeouts.
- MCP stdout, including bootstrap output, is JSON-RPC only.
- Agent-facing descriptions, server instructions, skills, and recovery messages
  must agree with actual defaults and advertised capabilities. Flag stale input
  names, nonexistent actions, missing parameter documentation, emojis, forced
  unrequested work, and screenshots required without an interactive desktop.
  Consent advice is not proof that server-side elicitation is implemented.
  Distinguish MCP input names from nested JSON keys, outputs, and CLI batch keys.
- Tests establish actual Excel state and returned fields, not only `Success`.
  Generated artifacts must change through their source; use code-derived counts
  rather than copied literals.

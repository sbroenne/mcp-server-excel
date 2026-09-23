---
applyTo: "**"
excludeAgent: "cloud-agent"
---

# ExcelMcp review priorities

Report high-confidence defects introduced by the PR, not style or unrelated
cleanup. Implementation guidance is excluded from review; retain these checks:

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
- Tests establish actual Excel state and returned fields, not only `Success`.
  Generated artifacts must change through their source; use code-derived counts
  rather than copied literals.

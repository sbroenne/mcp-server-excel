# Current architecture decisions

ADRs explain choices and tradeoffs, not commands or checklists. Actionable
contributor rules live in [AGENTS.md](../AGENTS.md) and its task-specific guides.
Use [CONTEXT.md](../CONTEXT.md) for the system map.

Read the records relevant to the architectural area being changed. Their
presence does not make them automatically loaded instructions.

| Record | Area | Status |
| --- | --- | --- |
| [ADR-001](ADR-001-TESTING-STRATEGY.md) | Real-Excel and Excel-independent testing | Current |
| [ADR-002](ADR-002-REAL-EXCEL-AUTOMATION.md) | Desktop Excel and typed interop | Current |
| [ADR-003](ADR-003-SHARED-SERVICE-SEPARATE-HOSTS.md) | Shared behavior, separate CLI/MCP hosts | Current |
| [ADR-004](ADR-004-GENERATED-COMMAND-CONTRACTS.md) | Generated operation contracts | Current |
| [ADR-005](ADR-005-SESSION-AND-FAILURE-OWNERSHIP.md) | Sessions, threading, failures, and cleanup | Current |
| [ADR-006](ADR-006-MCP-SDK-BOUNDARY.md) | MCP protocol ownership | Current |
| [ADR-007](ADR-007-LOCAL-ACCESS-AND-TELEMETRY.md) | Local access and privacy boundaries | Current |
| [ADR-008](ADR-008-GUIDANCE-SOURCE-OWNERSHIP.md) | Product and contributor guidance sources | Current |
| [ADR-009](ADR-009-COORDINATED-RELEASE-OUTPUTS.md) | Release and publication ownership | Current |

## Maintaining decisions

Create a record for a lasting architectural choice with meaningful alternatives,
not each operation, implementation detail, or coding rule. Explain the current
decision, reasons, tradeoffs, and consequences, then link its implementation
and authoritative guidance. Do not copy mutable defaults, checklists, or
procedures into the record.

Update a record when an approved decision changes. Keep this index aligned
with current records; Git retains older versions, so there is no ADR archive
or supersession chain. Proposed changes and unresolved questions belong in
the issue tracker until a decision is made. Do not invent past deliberations
or mark an unverified proposal as current.

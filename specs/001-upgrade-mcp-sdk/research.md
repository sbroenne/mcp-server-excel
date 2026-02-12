# Research: Upgrade MCP SDK to 0.5.0-preview.1

**Date**: 2025-12-13  
**Spec**: `specs/001-upgrade-mcp-sdk/spec.md`

---

## Research Summary

This document resolves all "NEEDS CLARIFICATION" or unknown items from the Technical Context by investigating SDK changelog, repository code, and .NET best-practice sources.

---

## 1. Compile-Breaking Deltas

### Decision

Bump the dependency and compile. Capture actual CS0619 (obsolete) and CS0117/CS0246 (removed) errors as the authoritative list.

### Rationale

Static analysis without bumping is incomplete; the compiler is the source of truth.

### Alternatives Considered

- Grepping for old API names in code. Rejected because: misses indirect usage and transitive issues.

---

## 2. Obsolete Enum Schema Types (MCP9001)

### Decision

Migrate all usages of obsolete "schema" types (if any) to the recommended equivalents in a single pass. Suppress nothing (FR-010).

### Rationale

Suppression leaves tech debt; upgrade path now is trivial because usage is not widespread.

### Alternatives Considered

- Warn-only (`<NoWarn>`). Rejected: violates zero-warnings constitution gate.

---

## 3. RequestOptions Migration

### Decision

Adopt `RequestOptions` bag pattern at all eligible call sites across MCP Server, Core, CLI, and tests (FR-011).

### Rationale

SDK deprecated positional parameters; uniform adoption prevents partial-upgrade fragmentation.

### Alternatives Considered

- Migrate MCP Server only. Rejected: test utilities still call into same layer; partial migration causes confusion.

---

## 4. `WithMeta` Adoption

### Decision

Adopt `WithMeta` for enriching tool responses (e.g., hints, suggested next actions) rather than embedding custom properties in JSON (FR-020).

### Rationale

`WithMeta` is the official extensibility mechanism; reduces custom JSON wrangling.

### Alternatives Considered

- Ignore for now. Rejected: user explicitly requested all new SDK features be adopted.

---

## 5. Console Application Best Practices (stdout / exit codes / shutdown)

### Decision

MCP server already logs to stderr (constitution-compliant). Enhancements:
- Ensure **no stdout output** after transport begins.
- Return exit code `1` on fatal error; `0` otherwise (FR-024, SC-015a).
- Observe cancellation token for graceful shutdown within 5 s (FR-026, SC-016).
- Verbosity configurable via env/config (FR-028, SC-017).

### Rationale

Aligns with .NET Generic Host guidance; ensures MCP transport stream is never polluted.

### Alternatives Considered

- Custom exit codes. Rejected: adds complexity; standard practice is 0/1.

---

## 6. New Attributes / Expanded Schema Attributes

### Decision

Audit MCP SDK for new or expanded attributes that improve tool/prompt metadata; adopt where applicable (FR-022, SC-013).

### Rationale

Attributes reduce boilerplate and produce cleaner schema.

### Alternatives Considered

- Defer until later upgrade. Rejected: user said "all new functionality"; cost is low.

---

## 7. URL-Mode Elicitation

### Decision

Out of scope (FR-019 deferred).

### Rationale

Project does not currently expose any remote prompts; minimal immediate value.

### Alternatives Considered

- Implement. Rejected: user explicitly said "I don't think we need this".

---

## 8. Error-Code Handling (ResourceNotFound −32002)

### Decision

Implement handling for `ResourceNotFound` error code from SDK where appropriate (FR-016, SC-010).

### Rationale

Ensures well-formed diagnostics for missing tool resources.

### Alternatives Considered

- Generic catch. Rejected: loses structured information.

---

## 9. UseStructuredContent for Tool Responses

### Decision

**Not adopted** for this project. Continue using serialized JSON in `TextContentBlock` responses.

### Background

SDK 0.5.0 introduces `UseStructuredContent` on `[McpServerTool]` attribute:
- When enabled, `Tool.OutputSchema` is populated with JSON Schema for the return type
- `CallToolResult.StructuredContent` contains typed JSON response (alongside `Content`)
- Return descriptions move from tool description into the schema

**Current approach (text-based JSON):**
```json
{
  "content": [{"type": "text", "text": "{\"success\": true, \"tables\": [...]}"}]
}
```

**With UseStructuredContent:**
```json
{
  "content": [{"type": "text", "text": "Operation completed"}],
  "structuredContent": {"success": true, "tables": [...]}
}
```

### Rationale

**Not suitable for action-based tool architecture:**

| Issue | Impact |
|-------|--------|
| **Action polymorphism** | Each action returns different result types (List → array, Create → single item, Delete → success flag). SDK expects ONE return type per tool. |
| **Return type mismatch** | Our tools return `Task<string>` with serialized JSON. UseStructuredContent expects actual typed objects. |
| **Significant refactoring** | 12 tools × 5-15 actions each = ~100+ result type classes to define |
| **Current JSON works** | LLMs parse our text JSON responses successfully |

**LLM Benefit Assessment:**

| Benefit | Assessment |
|---------|------------|
| Schema introspection | 🟡 Moderate - LLMs get schema upfront, but already handle our JSON well |
| Response validation | 🟡 Moderate - SDK validates against schema, but responses are consistent |
| Structured parsing | 🟢 Minor - `structuredContent` easier to parse, but clients handle text JSON |

### Alternatives Considered

1. **Split into single-action tools** - Would create 100+ tools, breaking clean API design. Rejected.
2. **Complex union types per tool** - Would require `OneOf` schemas for each action's return type. High complexity, low benefit. Rejected.
3. **Adopt for simple tools only** - Inconsistent API experience. Rejected.

### Future Consideration

UseStructuredContent is well-suited for single-purpose tools with consistent return types. If we ever split tools into individual operations (breaking change), this could be reconsidered.

---

## 10. Behavioral Hints (ReadOnly, Destructive, Idempotent, OpenWorld)

### Decision

**Not adopted** for this project due to action-based tool architecture.

### Background

SDK 0.5.0 adds behavioral hint properties on `[McpServerTool]`:
- `ReadOnly` - Tool doesn't modify environment
- `Destructive` - Tool can perform destructive updates
- `Idempotent` - Repeated calls have no additional effect
- `OpenWorld` - Tool interacts with external entities

### Rationale

These hints apply at **tool level**, but our tools have **mixed action behaviors**:

| Tool | Example Actions | ReadOnly? | Destructive? |
|------|-----------------|-----------|--------------|
| table | List | ✅ Yes | ❌ No |
| table | Delete | ❌ No | ✅ Yes |
| table | Create | ❌ No | ❌ No |

Setting these at tool level would be **misleading to LLMs** - they'd assume all actions share the same behavior.

### Alternatives Considered

- Set to most conservative value (Destructive=true). Rejected: defeats purpose of hints.
- Split into single-action tools. Rejected: excessive API surface (see #9).

---

## 11. IconSource and Visual Properties

### Decision

**Not adopted**. UI-focused feature not needed for CLI-based Excel automation.

### Rationale

Our tools are consumed programmatically by LLMs and automation scripts, not displayed in visual UIs.

---

## Open Questions

None remaining. All TBD items have been resolved through changelog analysis and user clarification.

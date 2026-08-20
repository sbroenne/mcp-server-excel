---
applyTo: "skills/shared/**/*.md,src/ExcelMcp.McpServer/Prompts/**/*.md"
excludeAgent: "code-review"
---

# Shared LLM guidance

`skills/shared` is the source of truth. Release builds copy it into both agent
skills and generated MCP prompts. Do not hand-edit generated prompt output.

Write for an agent that already understands Excel, JSON, and MCP. Include only
repository-specific information:

- Which ExcelMcp tool to choose when nearby tools overlap
- Non-obvious action differences and destructive behavior
- Session, batch, save, and refresh semantics
- Server-specific return shapes or Excel limitations
- Short recovery guidance for predictable errors
- Concrete parameter examples only when schema descriptions are insufficient

Do not duplicate enum catalogs, types, required/optional flags, generic Excel
tutorials, protocol syntax, or CLI help. Keep guidance concise and task-oriented.
Use plain text markers rather than emojis.

When changing shared guidance:

1. Edit the canonical file in `skills/shared`.
2. Build Release to regenerate prompts and skill references.
3. Check generated changes for parity, not as an independent source.
4. Run the relevant LLM evaluation only when discoverability or workflow choice
   changed.

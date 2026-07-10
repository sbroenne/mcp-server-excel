---
"Sbroenne.ExcelMcp.McpServer": patch
---

Fix JSON Schema array items format for Gemini API compatibility (#672)

Removes `nullable: true` from array nodes and adds explicit `type: string` fallback for C# `object` nodes. This prevents MCP clients from emitting missing types or union schemas that the strict Gemini API validator rejects.

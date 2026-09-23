---
applyTo: "skills/**/*.md,skills/templates/**/*.sbn,src/ExcelMcp.McpServer/Prompts/**,src/ExcelMcp.Build.Tasks/**/*.cs"
excludeAgent: "code-review"
---

# Agent guidance sources

| Content | Edit here |
|---------|-----------|
| Tool/parameter descriptions | Core interface XML docs/attributes; manual MCP metadata only where it owns the tool |
| Skill prose and selection rules | `skills/templates/SKILL.cli.sbn`, `SKILL.mcp.sbn` |
| Shared workflows and limitations | `skills/shared/*.md` |
| Skill rendering | `src/ExcelMcp.Build.Tasks/GenerateSkillFile.cs` |
| MCP prompt description overrides | `GenerateSkillPromptsClass` in the MCP `.csproj` |

Release builds generate both `skills/excel-*/SKILL.md` files from templates and
the Core manifest, copy shared references, and generate/embed MCP prompts.
Never edit those outputs or the extension's packaged skill copy.

Guidance should add only what an Excel-capable agent cannot infer from schemas:
tool disambiguation, server-specific semantics, pitfalls, and recovery.
No enum catalogs, generic Excel tutorials, duplicated CLI help, or emojis.
Rebuild Release after source edits; run affected evaluations when discovery or
workflow selection changes.

Generation pipeline and authoring procedure: `skills/README.md`.

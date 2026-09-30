---
applyTo: "skills/**/*.md,skills/templates/**/*.sbn,src/ExcelMcp.Build.Tasks/**/*.cs"
excludeAgent: "code-review"
---

# Agent guidance sources

| Content | Edit here |
|---------|-----------|
| Tool/parameter descriptions | Core interface XML docs/attributes; manual MCP metadata only where it owns the tool |
| Skill prose and selection rules | `skills/templates/SKILL.cli.sbn`, `SKILL.mcp.sbn` |
| Shared workflows and limitations | `skills/shared/*.md` |
| Skill rendering | `src/ExcelMcp.Build.Tasks/GenerateSkillFile.cs` |
| Minimal MCP server instructions | `Program.cs` in the MCP Server |

Release builds generate the Core manifest. Shared guides are not exposed as MCP prompts.
`scripts\Build-AgentSkills.ps1 -GenerateOnly` then generates complete skills under
`artifacts\generated-skills` from templates, authored `skills/assets`, and shared references.
Never edit those outputs or the extension's packaged skill copy.

Guidance should add only what an Excel-capable agent cannot infer from schemas:
tool disambiguation, server-specific semantics, pitfalls, and recovery.
No enum catalogs, generic Excel tutorials, duplicated CLI help, or emojis.
Rebuild Release after source edits; run affected evaluations when discovery or
workflow selection changes.
Do not require unnecessary formatting, Tables, questions, or presentation menus.
Discover existing state where useful; ask rather than guess when the intended
workbook, destructive change, or requested result is unclear.

Generation pipeline and authoring procedure: `skills/README.md`.

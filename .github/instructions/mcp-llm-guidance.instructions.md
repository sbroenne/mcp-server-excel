---
applyTo: "skills/**/*.md,skills/templates/**/*.sbn,src/ExcelMcp.Build.Tasks/**/*.cs,src/ExcelMcp.Core/Commands/**/*.cs,src/ExcelMcp.Generators.Mcp/**/*.cs,src/ExcelMcp.McpServer/**/*.cs"
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
Installed plugins and the published plugin repository are outputs too; local
source edits do not update installed skills or authorize publication.

Guidance should add only what an Excel-capable agent cannot infer from schemas:
tool disambiguation, server-specific semantics, pitfalls, and recovery.
No enum catalogs, generic Excel tutorials, duplicated CLI help, or emojis.
Rebuild Release after source edits; run affected evaluations when discovery or
workflow selection changes.
Do not require unnecessary formatting, Tables, questions, or presentation menus.
Discover existing state where useful; ask rather than guess when the intended
workbook, destructive change, or requested result is unclear.

Keep tool descriptions, parameter/action guidance, server instructions, both
skill templates, shared references, and returned recovery messages consistent.
MCP prose uses advertised snake_case inputs; CLI flags use kebab-case and batch
JSON uses Service camelCase. Do not rename nested JSON keys, enum values, output
fields, or external API identifiers when correcting top-level input names.

Restore the prior calculation mode after temporary changes, including failure
recovery. Screenshots depend on an interactive desktop; do not require them for
unattended jobs. Check actual create/load/refresh and save/close defaults.

The server currently exposes tools, not prompts or resources, and makes no MCP
elicitation requests. Instructions to obtain consent are client guidance, not
an enforced server confirmation dialog. Do not imply otherwise.

Core XML documentation is a build input to the MCP generator. Do not delete it
before downstream generation or accept description strings containing only
required/valid-action suffixes. Use existing discovery and generated-skill tests
to cover recurring guidance defects; do not add a parallel audit framework.

Generation pipeline and authoring procedure: `skills/README.md`.

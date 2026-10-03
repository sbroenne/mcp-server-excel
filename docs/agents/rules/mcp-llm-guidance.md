# Agent guidance sources

| Content | Edit here |
|---------|-----------|
| Tool/parameter descriptions | Core interface XML docs/attributes; manual MCP metadata only where it owns the tool |
| Generated command metadata | `src/ExcelMcp.Generators/ServiceRegistryGenerator.cs` and shared generator metadata |
| Skill prose and selection rules | `skills/excel-cli-report-formatting/SKILL.md`, `skills/excel-mcp-report-formatting/SKILL.md` |
| Shared workflows and limitations | `docs/reference/*.md` |
| Skill reference preparation | `scripts/Build-AgentSkills.ps1` |
| Minimal MCP server instructions | `Program.cs` in the MCP Server |

Release builds generate the Core manifest. Shared guides are not exposed as MCP prompts.
`scripts\Build-AgentSkills.ps1 -GenerateOnly` then generates complete skills under
`artifacts\generated-skills` from the two actual source skills and only the
canonical report-formatting reference. General documentation stays in `docs`.
Never edit those outputs or the extension's packaged skill copy.
Installed plugins and the published plugin repository are outputs too; local
source edits do not update installed skills or authorize publication.

Guidance should add only what an Excel-capable agent cannot infer from schemas:
tool disambiguation, server-specific semantics, pitfalls, and recovery.
No enum catalogs, generic Excel tutorials, duplicated CLI help, or emojis.
Keep server instructions minimal and task-focused; do not present optional
guides as required server instructions.
Rebuild Release after source edits; run affected evaluations when discovery or
workflow selection changes.
Do not require unnecessary formatting, Tables, questions, or presentation menus.
Discover existing state where useful; ask rather than guess when the intended
workbook, destructive change, or requested result is unclear.
When an operation fails after changing a workbook, do not assume it rolled back
or automatically undo/clean up the change. Use the reported partial-state
details to inspect the affected workbook, then decide whether cleanup is needed.

Keep tool descriptions, parameter/action guidance, server instructions, both
skill entries, documentation references, and returned recovery messages consistent.
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
required/valid-action suffixes. Verify emitted MCP descriptions through SDK
discovery, not only source strings. Use existing discovery and generated-skill
tests to cover recurring guidance defects; do not add a parallel audit framework.

Generation pipeline and authoring procedure: `docs/AGENT-SKILLS.md`.
